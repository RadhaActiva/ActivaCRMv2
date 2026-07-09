<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            width: 32%;
        }
        .auto-style2 {
            height: 29px;
        }
        .auto-style3 {
            width: 32%;
            height: 29px;
        }
        .auto-style4 {
            height: 27px;
        }
        .auto-style5 {
            width: 32%;
            height: 27px;
        }
    </style>
</head>
<%

set rs = server.CreateObject("adodb.recordset")

if request("md_code") <> "" then	  
sql = "SELECT md_id, md_code, md_desc, md_category, md_model, md_barcode, md_type, md_status, md_unitprice, md_brands, md_rcpprice, " & _
		"md_costprice, md_averageecost, md_attr_code, md_stock_uom, md_unitprice1, md_unitprice2, md_unitprice3, md_unitprice4,  " & _
		"md_logby, md_logdate, md_color, md_size " & _
		"FROM tblmodel WHERE md_code = '" & request("md_code") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			md_id = rs("md_id") 
			md_code = rs("md_code") 
			md_desc = rs("md_desc") 
			md_category = rs("md_category") 
			md_model = rs("md_model") 
			md_barcode = rs("md_barcode") 
			md_type = rs("md_type") 
			md_status = rs("md_status") 
			md_unitprice = chknumber2(rs("md_unitprice"))
			md_brands = rs("md_brands") 
			md_rcpprice = chknumber2(rs("md_rcpprice")) 
			md_costprice = chknumber2(rs("md_costprice"))
			md_averageecost = chknumber2(rs("md_averageecost")) 
			md_attr_code = rs("md_attr_code")
			md_stock_uom = rs("md_stock_uom")
			md_unitprice1 = chknumber2(rs("md_unitprice1"))
			md_unitprice2 = chknumber2(rs("md_unitprice2"))
			md_unitprice3 = chknumber2(rs("md_unitprice3"))
			md_unitprice4 = chknumber2(rs("md_unitprice4"))
		
			md_logby = rs("md_logby")
			md_logdate = rs("md_logdate")		
			md_color = rs("md_color")
			md_size = rs("md_size")			
		End If
		rs.Close
	  stype = "editStock"	
	  actionname = "Save" 
end if

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Update </font>Stock Master</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#666666">
                   
                   <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                    <tr>
                      <td width="26%" bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"> Stock Code*</font></font></strong></td>
                      <td class="auto-style1"><strong><%=md_code%></strong>
                        <input type="hidden" name="md_code" id="md_code" value="<%=md_code%>"/>
                        <input type="hidden" name="md_id" id="md_id" value="<%=md_id%>" /></td>
                      <td width="31%" bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Latest Purchase Price, RM</strong></font></strong></td>
                      <td width="25%"><strong>
                      <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                          <input name="md_costprice" type="text" id="md_costprice" value="<%=md_costprice%>" size="15" maxlength="10" />					
                      <%else%>
                      Restricted View
                      <%end if%></strong>
                      </td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF">Status</font></strong></td>
                      <td valign="top" class="auto-style1"><strong>
                        <select name="md_status" id="md_status">
                          <option value="N" <%if md_status = "Y" then response.write " selected"%>>N</option>
                          <option value="Y" <%if md_status = "Y" then response.write " selected"%>>Y</option>
                        </select>
                      </strong></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Average Cost Price, RM</strong></font></strong></td>
                      <td><strong>
					  <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                          <input name="md_averageecost" style="background-color: #cccccc;" type="text" id="md_averageecost" value="<%=md_averageecost%>" size="15" maxlength="10" readonly/>					 
                      <%else%>
                      Restricted View
                      <%end if%>
                      </strong></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF">Description</font></strong></td>
                      <td valign="top" class="auto-style1">
                          <textarea name="md_desc" cols="50" rows="3" id="md_desc"><%=md_desc%></textarea> </td>  
                          <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>RCP Price (West Malaysia) - RM</strong></font></strong></td>
                      <td><input name="md_rcpprice" type="text" id="md_rcpprice" value="<%=md_rcpprice%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td class="auto-style2" bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Category</strong></font></td>
                      <td valign="top" class="auto-style3">
                      <select name="md_category" id="md_category">
                          <option value="Labour" <%if md_category="Labour" then response.write " selected"%>>Labour</option>
                          <option value="FAN : CEILING FAN" <%if md_category="FAN : CEILING FAN" then response.write " selected"%>>FAN : CEILING FAN</option>
                          <option value="WHEAT : WATER HEATER" <%if md_category="WHEAT : WATER HEATER" then response.write " selected"%>>WHEAT : WATER HEATER</option>
                          <option value="Component" <%if md_category="Component" then response.write " selected"%>>Component</option>
                          <option value="SWHEA : STORAGE WATER HEATER" <%if md_category="SWHEA : STORAGE WATER HEATER" then response.write " selected"%>>SWHEA : STORAGE WATER HEATER</option>
                          <option value="Parts" <%if md_category="Parts" then response.write " selected"%>>Parts</option>
                          <option value="Service" <%if md_category="Service" then response.write " selected"%>>Service</option>    
                          <option value="SF" <%if md_category="SF" then response.write " selected"%>>SF</option>                          
                          <option value="GENERAL FAN" <%if md_category="GENERAL FAN" then response.write " selected"%>>GENERAL FAN</option>                          
                          
                        </select>
                      </td>
                      <td class="auto-style2" bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Selling Price1 (West Malaysia) - RM</strong></font></strong></td>
                      <td class="auto-style2"><input name="md_unitprice1" type="text" id="md_unitprice1" value="<%=md_unitprice1%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Brand</strong></font></td>
                      <td valign="top"><input name="md_brands" type="text" id="md_brands" value="<%=md_brands%>" size="50" maxlength="50" /></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Selling Price2 (East Malaysia) - RM</strong></font></strong></td>
                      <td><input name="md_unitprice2" type="text" id="md_unitprice2" value="<%=md_unitprice2%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Model </strong></font></td>
                      <td valign="top"><input name="md_model" type="text" id="md_model" value="<%=md_model%>" size="50" maxlength="50" /></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Selling Price3 (Singapore) - RM</strong></font></strong></td>
                      <td><input name="md_unitprice3" type="text" id="md_unitprice3" value="<%=md_unitprice3%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Type </strong></font></td>
                      <td valign="top" class="auto-style1"><font size="2">
                        <select name="md_type" id="md_type">
                          <option value="CF" <%if md_type="CF" then response.write " selected"%>>CF</option>
                          <option value="WH" <%if md_type="WH" then response.write " selected"%>>WH</option>
                          <option value="Service" <%if md_type="Service" then response.write " selected"%>>Service</option>
                        </select>
                      </font></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF">Selling Price4, RM</font></strong></td>
                      <td><input name="md_unitprice4" type="text" id="md_unitprice4" value="<%=md_unitprice4%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>UOM</strong></font></td>
                      <td valign="top" class="auto-style1"><strong><%=md_stock_uom%></strong></td>
                    </tr>
                    <tr>
                      <td class="auto-style4" bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Color</strong></font></td>
                      <td valign="top" class="auto-style5"><strong><%=md_color%>  | Size: <%=md_size%></strong></td>                     
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Barcode</strong></font></td>
                      <td class="auto-style1"><strong><%=md_barcode%></strong></td>                    
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Last Updated by </strong></font></td>
                      <td class="auto-style1"><strong><%=md_logby%></strong> @ <strong><%=chkdatetime(md_logdate)%></strong></td>                      
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>                                         
                      <td colspan="2"><a href="javascript:popup('rm_rpt_inventory_productgroup_detail.asp?stk_itm_code=<%=md_code%>','cb18','scrollbars=yes,resizable=yes,width=500,height=500')"><strong>View Stock Movement List</strong></a></td>
                      <td></td>  
                        </tr>
                    <tr>
                      <td colspan="4" align="right">
                    
                      <input type="submit" name="button" id="button" value="<%=actionname%>" />
                   
                      </td>
                    </tr>
                    </form>
                    
                    <tr>
                      <td colspan="4">&nbsp;</td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="4" bgcolor="#CCCCCC"><strong>Store Stock</strong></td>
                    </tr>
                    <tr>
                      <td colspan="4"><table width="100%" border="1" cellspacing="0" cellpadding="3">
                        <tr>
                          <th width="6%" align="center" bgcolor="#475387" scope="row"> <font size="2" color="#FFFFFF">No.</font></th>
                          <th width="11%" align="left" bgcolor="#475387" scope="row"><font size="2" color="#FFFFFF">Store Code</font></th>
                          <th width="18%" align="left" bgcolor="#475387" scope="row"><font size="2" color="#FFFFFF">Store Name</font></th>
                          <th width="17%" align="left" bgcolor="#475387" scope="row"><font size="2" color="#FFFFFF">Person In Charge</font></th>
                          <th width="13%" align="left" bgcolor="#475387" scope="row"><font size="2" color="#FFFFFF">Tel.</font></th>
                          <th width="23%" align="left" bgcolor="#475387" scope="row"><font size="2" color="#FFFFFF">Remark</font></th>
                          <td width="12%" align="center" bgcolor="#475387"><strong><font size="2" color="#FFFFFF"><strong>Current Qty</strong></font></strong></td>
                          </tr>
                          
<%
i = 1
wst_itm_current_qty = 0	   
sql = "Select tblstocktran.stk_reference as wh_code, " & _
		"(select top 1 tblwarehouse.wh_name from tblwarehouse where tblwarehouse.wh_code=tblstocktran.stk_reference) as wh_name, " & _
		"(select top 1 tblwarehouse.wh_contact_person from tblwarehouse where tblwarehouse.wh_code=tblstocktran.stk_reference) as wh_contact_person," & _
		"(select top 1 tblwarehouse.wh_remark from tblwarehouse where tblwarehouse.wh_code=tblstocktran.stk_reference) as wh_remark, " & _
		"(select top 1 tblwarehouse.wh_tel from tblwarehouse where tblwarehouse.wh_code=tblstocktran.stk_reference) as wh_tel, " & _
		"sum(tblstocktran.stk_qty) totalqty, round(sum(tblstocktran.stk_qty*tblmodel.md_averageecost),2) as totalvalue " & _
		"from tblstocktran inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code " & _
		"where tblstocktran.stk_reference <> '0' and tblmodel.md_code='" & md_code & "' group by tblstocktran.stk_reference " & _
		"order by totalqty desc "  
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
	  row = 100
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
                        <tr>
                          <th align="center" scope="row"><%=count%>.</th>
                          <th align="left" scope="row"><%=rs("wh_code")%></th>
                          <th align="left" scope="row"><%=rs("wh_name")%></th>
                          <th align="left" scope="row"><%=rs("wh_contact_person")%></th>
                          <th align="left" scope="row"><%=rs("wh_tel")%></th>
                          <th align="left" scope="row"><%=rs("wh_remark")%></th>
                          <td align="center"><strong><%=rs("totalqty")%></strong></td>
                          </tr>
 <%
 wst_itm_current_qty = wst_itm_current_qty + rs("totalqty")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>            
                          
                        <tr>
                          <th colspan="6" align="right" bgcolor="#CCCCCC" scope="row">Total</th>
                          <td align="center" bgcolor="#CCCCCC"><strong><%=wst_itm_current_qty%></strong></td>
                          </tr>
                      </table></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
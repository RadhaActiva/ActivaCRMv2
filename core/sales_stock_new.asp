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
        .auto-style6 {
            color: #CC0000;
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
                        <td class="titleblue1"><div align="left"><span class="auto-style6">View </span>Stock Master</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#666666">
                   
                   <form name="formorder" method="post">
                    <tr>
                      <td width="26%" bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"> Stock Code*</font></font></strong></td>
                      <td class="auto-style1"><strong><%=md_code%></strong>
                        <input type="hidden" name="md_code" id="md_code" value="<%=md_code%>"/>
                        <input type="hidden" name="md_id" id="md_id" value="<%=md_id%>" /></td>
                      <td width="31%" bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Latest Purchase Price, RM</strong></font></strong></td>
                      <td width="25%"><strong>
                      <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                          <input name="md_costprice" type="text" id="md_costprice" value="******" size="15" maxlength="10" />					
                      <%else%>
                      Restricted View
                      <%end if%></strong>
                      </td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF">Status</font></strong></td>
                      <td valign="top" class="auto-style1"><strong> <%=md_status%></strong></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Average Cost Price, RM</strong></font></strong></td>
                      <td><strong>
					  <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                          <input name="md_averageecost" style="background-color: #cccccc;" type="text" id="md_averageecost" value="*****" size="15" maxlength="10" readonly/>					 
                      <%else%>
                      Restricted View
                      <%end if%>
                      </strong></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF">Description</font></strong></td>
                      <td valign="top" class="auto-style1">
                          <textarea name="md_desc" cols="50" rows="3" id="md_desc" readonly><%=md_desc%></textarea> </td>  
                          <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>RCP Price (West Malaysia) - RM</strong></font></strong>
                      </td>
                      <td><%=md_rcpprice%></td>
                    </tr>
                    <tr>
                      <td class="auto-style2" bgcolor="#CD6155"><font size="2" color="#FFFFFF" readonly><strong>Category</strong></font></td>
                      <td valign="top" class="auto-style3"><%=md_category%></td>
                      <td class="auto-style2" bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Selling Price1 (West Malaysia) - RM</strong></font></strong></td>
                      -<td class="auto-style2">Restricted View</td>
                        
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Brand</strong></font></td>
                      <td valign="top"><%=md_brands%></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Selling Price2 (East Malaysia) - RM</strong></font></strong></td>
                      <td>Restricted View</td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Model </strong></font></td>
                      <td valign="top"><%=md_model%></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Selling Price3 (Singapore) - RM</strong></font></strong></td>
                      <td>Restricted View</td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Type </strong></font></td>
                      <td valign="top" class="auto-style1" ><%= md_type%><font size="2" readonly>               
                      </font></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Selling Price4, RM</strong></font></strong></td>
                      <td>Restricted View</td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>UOM</strong></font></td>
                      <td valign="top" class="auto-style1"> <strong><%=md_stock_uom%></strong></td>
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
                      <td colspan="2"></td>
                      <td></td>  
                        </tr>
                    <tr>
                      <td colspan="4" align="right">
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
		"where tblstocktran.stk_reference ='W1' and tblmodel.md_code='" & md_code & "' group by tblstocktran.stk_reference " & _
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
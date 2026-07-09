<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            height: 24px;
        }
        .auto-style2 {
            height: 29px;
        }
    </style>
</head>
<%

actionname = "Save" 
stype = "AddStock"	

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Create </font>Stock Master</div></td>
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
                      <td width="26%" bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"> Stock Code<font color="#FF0000">*</font></font></strong></td>
                      <td width="28%"><strong><%=md_code%></strong>
                        <input type="text" name="md_code" id="md_code" value="<%=md_code%>"/>
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
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Status Active</strong></font></td>
                      <td valign="top"><strong>
                        <select name="md_status" id="md_status">
                          <option value="N" <%if md_status = "Y" then response.write " selected"%>>N</option>
                          <option value="Y" <%if md_status = "Y" then response.write " selected"%>>Y</option>
                        </select>
                      </strong></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Average Cost Price, RM</strong></font></strong></td>
                      <td><strong>
					  <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                          <input name="md_averageecost" type="text" id="md_averageecost" value="<%=md_averageecost%>" size="15" maxlength="10" readonly />
                      <%else%>
                      Restricted View
                      <%end if%>
                      </strong></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF">Description</font></strong></td>
                      <td valign="top"><strong>
                          <textarea name="md_desc" cols="50" rows="3" id="md_desc"><%=md_desc%></textarea>                          
                          </strong></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF"><strong>Retail Price (West Malaysia) - RM</strong></font></strong>
                      </td>
                      <td><input name="md_rcpprice" type="text" id="md_rcpprice" value="<%=md_rcpprice%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Category</strong></font></td>
                      <td valign="top">
                      <select name="md_category" id="md_category">
                          <option value="FAN : CEILING FAN" <%if md_category="FAN : CEILING FAN" then response.write " selected"%>>FAN : CEILING FAN</option>
                          <option value="WHEAT : WATER HEATER" <%if md_category="WHEAT : WATER HEATER" then response.write " selected"%>>WHEAT : WATER HEATER</option>
                          <option value="Component" <%if md_category="Component" then response.write " selected"%>>Component</option>
                          <option value="SWHEA : STORAGE WATER HEATER" <%if md_category="SWHEA : STORAGE WATER HEATER" then response.write " selected"%>>SWHEA : STORAGE WATER HEATER</option>
                          <option value="Parts" <%if md_category="Parts" then response.write " selected"%>>Parts</option>
                          <option value="Labour" <%if md_category="Labour" then response.write " selected"%>>Labour</option>
                          <option value="Service" <%if md_category="Service" then response.write " selected"%>>Service</option>
                        </select>
                      </td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF">Selling Price1 (West Malaysia) - RM</font></strong></td>
                      <td><input name="md_unitprice1" type="text" id="md_unitprice1" value="<%=md_unitprice1%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Brand</strong></font></td>
                      <td valign="top"><strong><input name="md_brands" type="text" id="md_brands" value="<%=md_brands%>" size="40" maxlength="40" /></strong></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF">Selling Price2 (East Malaysia) - RM</font></strong></td>
                      <td><input name="md_unitprice2" type="text" id="md_unitprice2" value="<%=md_unitprice2%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Model </strong></font></td>
                      <td valign="top"><strong><input name="md_model" type="text" id="md_model" value="<%=md_model%>" size="40" maxlength="40" /></strong></td>
                      <td bgcolor="#CD6155"><strong><font size="2" color="#FFFFFF">Selling Price3 (Singapore) - RM</font></strong></td>
                      <td><input name="md_unitprice3" type="text" id="md_unitprice3" value="<%=md_unitprice3%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Type </strong></font></td>
                      <td valign="top"><font size="2">
                        <select name="md_type" id="md_type">
                          <option value="WH" <%if md_type="WH" then response.write " selected"%>>WH</option>
                          <option value="CF" <%if md_type="CF" then response.write " selected"%>>CF</option>
                          <option value="Service" <%if md_type="Service" then response.write " selected"%>>Service</option>
                        </select>
                      </font></td>
                      <td bgcolor="#CD6155"<strong><strong><font size="2" color="#FFFFFF">Selling Price4, RM</font></strong></td>
                      <td><input name="md_unitprice4" type="text" id="md_unitprice4" value="<%=md_unitprice4%>" size="15" maxlength="10" /></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>UOM</strong></font></td>
                      <td valign="top">
                          <select name="md_stock_uom" id="md_stock_uom">
                          <option value="PCS" <%if md_type="PCS" then response.write " selected"%>>PCS</option>
                          <option value="TRIP" <%if md_type="TRIP" then response.write " selected"%>>TRIP</option>
                        </select>
                      <td></td>
                      <td></td>
                    </tr>
                    <tr>
                      <td class="auto-style2" bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Color</strong></font></td>
                      <td valign="top" class="auto-style2"><input name="md_color" type="text" id="md_color" value="<%=md_color%>" size="30" maxlength="30" />
                      <td class="auto-style2"></td>                      
                    </tr>
                     <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Size</strong></font></td>
                      <td valign="top"><input name="md_size" type="text" id="md_size" value="<%=md_size%>" size="30" maxlength="30" /></td>
                      <td></td>
                      <td></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Barcode</strong></font></td>
                      <td valign="top"><input name="md_barcode" type="text" id="md_barcode" value="<%=md_barcode%>" size="40" maxlength="40" /></td>
                      <td></td>
                    </tr>
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF"><strong>Group Type</strong></font></td>
                      <td><strong><input name="md_group_type" type="text" id="md_group_type" value="<%=md_group_type%>" size="40" maxlength="40" /></strong></td>
                      <td></td>
                      <td></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    
                    <tr>
                      <td bgcolor="#CD6155"><font size="2" color="#FFFFFF">Last Updated by </font></td>
                      <td><strong><%=md_logby%></strong> @ <strong><%=chkdatetime(md_logdate)%></strong></td>
                      <td></td>
                      <td></td>
                    </tr>
                    
                    <tr>
                      <td colspan="4" align="right">
                      <input type="submit" name="button" id="button" value="Create Stock" />     
                      </td>
                    </tr>
                    </form>
                       
                    <tr>
                      <td colspan="4" class="auto-style1"></td>
                    </tr>
                    <tr>
                      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                        <tr>                          
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
<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=salesanalysis_detail_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->
<table width="100%" border="1" cellpadding="4" cellspacing="0">
 <tr>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Invoice No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Date</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer Code<br />
                        </span></strong></font><font color="#FFFFFF"><strong><span><br />
                      </span></strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Customer Name</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Code</font></strong></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Type</font></strong></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Description</font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong> Inv<br />
                      </strong></font><font color="#FFFFFF"><strong>Qty</strong></font></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Inv<br /> 
                      Amt</span></strong></font></td>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Gst<br />
                      Amt</font></strong></td>
                      <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Part <br />
                        Sales Amt</font></strong></td>
                      <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Labour</font><font color="#FFFFFF"> <br />
                        Amt</font></strong></td>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Transport <br />
                       Amt</font></strong></td>
                      <!--<td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Part Cost</font></strong></td>-->					  
					  <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Average Cost Price(Per Item)</font></strong></td>
					  <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Average Cost Price</font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Gross Profit</span></strong></font></td>
                      <td height="30" align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>GP %</span></strong></font></td>
                    </tr>
  <%
i = 1  
sql2 = request.Cookies("GAPS")("sqlexcel")  
'response.write sql2
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
invd_costtotal=0
invd_grossamt=0
while not rs.eof 
%>
  <tr>
                      <td height="40" align="center" nowrap="nowrap"><%=i%>.</td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><a href="rm_invoice_new.asp?inv_no=<%=rs("inv_no")%>" target="_blank"><%=rs("inv_no")%></a></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("inv_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=(rs("inv_status"))%></td>
                      <td align="left"><%=rs("inv_cust_code")%></td>
                      <td align="left"><%=rs("inv_cust_name")%></td>
                      <td align="left" nowrap="nowrap"><%=rs("invd_partcode")%></td>
                      <td align="left"><%=rs("invd_parttype")%></td>
                      <td align="left"><%=rs("invd_desc")%></td>
                      <td align="right" nowrap="nowrap"><strong> <%=rs("invd_qty")%></strong></td>
                      <td align="right"><strong><%=chknumber2(rs("invd_subtotal"))%></strong></td>
                      <td align="right"><strong>0.00</strong></td>
                     <!-- <td align="right"><strong><%=chknumber2(rs("invd_subtotal")*0.0566037735849057)%></strong></td>-->
                      <td align="right"><strong><%if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal"))%></strong></td>
                      <td align="right"><strong><%if rs("invd_parttype") = "Labour" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal"))%></strong></td>
                      <td align="right"><strong><%if rs("invd_parttype") = "Transport" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal"))%></strong></td>

                      <!--<td align="right"><strong><%if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057))%></strong></td>
                      <td align="right"><strong><%if rs("invd_parttype") = "Labour" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057))%></strong></td>
                      <td align="right"><strong><%if rs("invd_parttype") = "Transport" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057))%></strong></td>-->
                      <td align="center">
					  <strong><!--check total excel-->
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(rs("md_averageecost"))%>                         
                        <%end if%>
                      </strong>
					  </td>
					  <td align="center">
					  <!--<strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(rs("invd_itemcost")*rs("invd_qty"))%>
                        <%end if%>
                      </strong>-->
					  <strong><!--check total excel-->
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(rs("md_averageecost")*rs("invd_qty"))%>
                        <%end if%>
                      </strong>
					  </td>
                      <td align="right"><strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <!--<%if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057)-(rs("invd_itemcost")*rs("invd_qty")))%>-->
                          <%if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal")-(rs("invd_itemcost")*rs("invd_qty")))%>
                        <%end if%>
                      </strong></td>
                      <td align="right" nowrap="nowrap"><strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                         <%if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then 
						     'response.write chknumber2((rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057)-(rs("invd_itemcost")*rs("invd_qty")))/(rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057))*100) 
                             response.write chknumber2((rs("invd_subtotal")-(rs("invd_subtotal"))-(rs("invd_itemcost")*rs("invd_qty")))/(rs("invd_subtotal"))*100) 
						   else
						      response.write "0.00"
						   end if	  
							  %>
                        <%end if%>
                      %</strong></td>
                    </tr>
  <%
invd_qty = invd_qty + ccur(rs("invd_qty"))
invd_subtotal = invd_subtotal + ccur(rs("invd_subtotal"))
inv_gstAmt = inv_gstAmt + ccur((rs("invd_subtotal")*0.0566037735849057))

if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then 
'invd_parts = invd_parts + ccur(rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057))
invd_parts = invd_parts + ccur(rs("invd_subtotal"))
end if

if rs("invd_parttype") = "Labour" and rs("invd_subtotal") > 0 then 
'invd_Labour = invd_Labour + ccur(rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057))
invd_Labour = invd_Labour + ccur(rs("invd_subtotal"))
end if

if rs("invd_parttype") = "Transport" and rs("invd_subtotal") > 0 then 
'invd_Transport = invd_Transport + ccur(rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057))
invd_Transport = invd_Transport + ccur(rs("invd_subtotal"))
end if

invd_costtotal = invd_costtotal + chknumber2(rs("md_averageecost"))
'invd_costtotal = invd_costtotal + ccur(rs("invd_itemcost")*rs("invd_qty"))
'invd_costtotal = invd_costtotal + ccur(rs("invd_itemcost"))

'if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then 
'invd_grossamt = invd_grossamt + ccur(rs("invd_subtotal")-(rs("invd_subtotal")*0.0566037735849057)-(rs("invd_itemcost")*rs("invd_qty")))
invd_grossamt = invd_grossamt + chknumber2(rs("md_averageecost")*rs("invd_qty"))


'end if

count = count + 1 
i = i + 1
rs.MoveNext
wend
rs.Close

%>
   <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="9" align="right" bgcolor="#999999"><strong>Total</strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=(invd_qty)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(invd_subtotal)%></strong></td>
                      <td height="40" align="right" bgcolor="#999999"><strong>0.00</strong></td>
                      <!--<td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(inv_gstAmt)%></strong></td>-->
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(invd_parts)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(invd_Labour)%></strong></td>
                      <td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(invd_Transport)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(invd_costtotal)%>
                        <%end if%>
                      </strong></td>
                      <td align="right" bgcolor="#999999"><strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(invd_grossamt)%>
                        <%end if%>
                      </strong></td>
                      <td align="right" bgcolor="#999999"><strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2((invd_grossamt/invd_parts)*100)%>
                        <%end if%>
                      %</strong></td>
                    </tr>
</table>

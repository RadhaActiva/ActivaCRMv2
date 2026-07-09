<!-- #include file="database/datastore.asp" -->
<%  
job_date_to = request("job_date_to")
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=Inventory_Assessment_Report_" & job_date_to & ".xls"


sql10="select * from tblrpr_new_pnl where pnl_date_ending = '" & job_date_to & "' and pnl_logby='" & Request.Cookies("GAPS")("sloginid") & "'" 
set rs10 = server.CreateObject("adodb.recordset")
rs10.ActiveConnection = strconnect
rs10.Source = sql10
rs10.CursorLocation  = 3
rs10.Open

if not rs10.eof then
    total_sales_parts=rs10("pnl_total_sales_parts")
    spare_parts_over_wrty=rs10("pnl_spare_parts_over_wrty")
    spare_parts_under_wrty=rs10("pnl_spare_parts_under_wrty")
    total_spare_cost=rs10("pnl_total_spare_cost")
    gross_profit=rs10("pnl_gross_profit")
    percentage_profit=rs10("pnl_percentage_profit")
    opening_stock=rs10("pnl_opening_stock")
    total_stock_purchases=rs10("pnl_total_stock_purchases")
    total_spare_parts_stock=rs10("pnl_total_spare_parts_stock")
    closing_stock_figure=rs10("pnl_closing_stock_figure")
    FTCRM_closing_stock=rs10("pnl_FTCRM_closing_stock")
    stock_value_diff=rs10("pnl_stock_value_diff")
    total_stock_out=rs10("pnl_total_stock_out")
    total_stock_adjust=rs10("pnl_total_stock_adjust")
    stock_cn=rs10("pnl_credit_notes")
    spare_parts_over_wrty_noinv=rs10("pnl_spare_parts_over_wrty_notinv")
    spare_parts_do=rs10("pnl_spare_parts_do")
    spare_parts_err=rs10("pnl_error_job")
end if    

%>

<!-- #include file="database/datastore.asp" -->

<td width="30%"><nowrap><h3>RIEGEN MARKETING SDN BHD</h3></nowrap></td>
<tr></tr>
<td width="30%"><nowrap><h3>MONTHLY INVENTORY ASSESSMENT REPORT- PERIOD ENDING <%=job_date_to%></h3></nowrap>
 <table width="60%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <br>
                        <br>
                        <tr><td></td><td></td><td></td><td style="text-align:center"><b>MYR</b></td></tr>
                        <tr>
                        <td></td>
                        <td width="30%"><nowrap>Sales Spare Parts</nowrap></td>
                        <td width="20%"></td>
                        <td><%=formatnumber(total_sales_parts)%></td>
                        <td></td>
                        </tr>
                        <P></P>
                        <tr><td width="104" height="5" align="right"></td></tr>
                        <tr><td></td><td width="15%"><nowrap><U>LESS : Cost of Sales</U></nowrap></td></tr>
                        <tr><td></td><td width="15%">Total Spare Part Costs - As Below</td>
                        <td></td>
                        <td><%=formatnumber(total_spare_cost)%></td></tr>
                        <tr><td width="104" height="5" align="right"></td></tr>
                        <td></td>
                        <td></td>
                        <td></td>
                        <tr><td></td><td valign="top" bgcolor="#F3F3F3" <strong><font color="#0000FF"><b>Gross Profit</b></font></strong></td><td></td>
                        <td bgcolor="#F3F3F3"><u><%=formatnumber(gross_profit)%></u></td></tr>
                        <tr><td></td><td valign="top">% Of Gross Profit</td><td></td>
                        <td><%=chknumber2(percentage_profit)%>%</td></tr>
                        <td></td>
                        <tr><td></td><td width="60%"><strong>Summary Stock</strong></td></tr>
                        <tr><td></td><td></td><td></td><td></td><td style="text-align:center"><b>MYR</b></td></tr>
                        <tr><td></td><td>Opening Stock</td>
                        <td></td><td></td><td><%=formatnumber(opening_stock)%></td>
                        <tr><td></td><td>ADD: Stock Purchases</td>
                        <td></td><td></td><td><%=formatnumber(total_stock_purchases)%></td>
                        <tr><td></td><td>ADD : Credit Notes</td>
                        <td></td><td></td><td><%=formatnumber(stock_cn)%></td>
                        <!--<tr><td width="104" height="10" align="right"></td></tr>-->

                        <tr><td></td><td width="60%">LESS : Stock-Out</td>
                        <td></td><td></td><td><u><%=formatnumber(total_stock_out,,,-1)%></u></td></tr>
                        <tr><td></td><td>(+/-): Stock Adjustments</td>
                        <td></td><td></td><td><%=formatnumber(total_stock_adjust)%></td>
                        <tr><td width="104" height="10" align="right"></td></tr>

                        <tr><td></td><td width="60%">Total Spare Parts Stock</td>
                        <td></td><td></td><td><u><%=formatnumber(total_spare_parts_stock)%></u></td></tr>
                        <tr><td width="104" height="5" align="right"></td></tr>
                        <tr><td></td><td width="60%">LESS: Cost of Sales - Spare Parts O/Warranty</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_over_wrty,,,-1)%></td></tr>
                        <tr><td></td><td width="60%">LESS: Cost of Sales - Spare Parts O/Warranty (Not Inv)</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_over_wrty_noinv,,,-1)%></td></tr>
                        <tr><td></td><td width="60%">LESS: Spare Part Costs (W/H & C/F)-U/Warranty</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_under_wrty,,,-1)%></td></tr>
                        <tr><td></td><td width="60%">LESS: DO Spare Part Costs</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_do,,,-1)%></td></tr>
                        <tr><td></td><td width="60%">LESS: MISC (Error Jobs)</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_err,,,-1)%></td></tr>
                        <tr><td width="104" height="5" align="right"></td></tr>
                        <tr><td></td><td width="60%">Closing Stock as at <%=job_date_to%></td>
                        <td></td><td></td><td><%=formatnumber(closing_stock_figure)%></td></tr>
                        <tr><td width="104" height="5" align="right"></td></tr>
                        <tr><td></td><td width="60%">As per FTCRM closing stock <%=job_date_to%> figure shows </td>
                        <td></td><td></td><td><u><%=formatnumber(FTCRM_closing_stock)%></u></td></tr>
                        <tr><td width="104" height="10" align="right"></td></tr>
                        <%if stock_value_diff < 0 then %>
                        <tr><td></td><td width="60%"><font color="#FF0000"><strong><b>STOCK VALUE DIFFERENCES</b></strong></font></td>
                        <td></td><td></td><td><b><u><font color="#FF0000"><%=formatnumber(stock_value_diff)%></u></b></font></td></tr>
                        <%else%>>
                        <tr><td></td><td width="60%"><font color="#0000FF"><strong><b>STOCK VALUE DIFFERENCES</b></strong></font></td>
                        <td></td><td></td><td><b><u><font color="#0000FF"><%=formatnumber(stock_value_diff)%></u></b></td></font></tr>
                        <%end if%> 
                  </table>

<%
If Err.Number <> 0 Then
  Response.Write (Err.Description)   
  Response.End 
End If
%>


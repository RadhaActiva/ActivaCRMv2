<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = request("job_actual_wrty_status")

if ordertype = "" then 
   ordertype = "desc"
end if

if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

if request("jobmonth") <> "" then
   jobmonth = request("jobmonth")
else
   jobmonth = month(date())
end if

if request("jobyear") <> "" then
   jobyear = request("jobyear")
else
   jobyear = year(date())
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Inventory</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1"> Stock-In and Stock-Out Report                          
                        <label for="select"></label></td>
                      <td width="20%" align="center" class="titlegrey1"><img src="images/excel.jpg" width="57" height="21" /></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="action_report.asp?type=monthcommisionIC">
                    <select name="jobyear" id="jobyear">
                      <option value="2016"<%if jobyear="2016" then response.write " selected"%>>2016</option>
                    </select>
                    <select name="jobmonth" id="jobmonth">
                      <option value="1" <%if jobmonth="1" then response.write " selected"%>>Jan</option>
                      <option value="2" <%if jobmonth="2" then response.write " selected"%>>Feb</option>
                      <option value="3" <%if jobmonth="3" then response.write " selected"%>>Mar</option>
                      <option value="4" <%if jobmonth="4" then response.write " selected"%>>Apr</option>
                      <option value="5" <%if jobmonth="5" then response.write " selected"%>>May</option>
                      <option value="6" <%if jobmonth="6" then response.write " selected"%>>Jun</option>
                      <option value="7" <%if jobmonth="7" then response.write " selected"%>>Jul</option>
                      <option value="8" <%if jobmonth="8" then response.write " selected"%>>Aug</option>
                      <option value="9" <%if jobmonth="9" then response.write " selected"%>>Sep</option>
                      <option value="10" <%if jobmonth="10" then response.write " selected"%>>Oct</option>
                      <option value="11" <%if jobmonth="11" then response.write " selected"%>>Nov</option>
                      <option value="12" <%if jobmonth="12" then response.write " selected"%>>Dec</option>
                    </select>
                    <span class="titlegrey1">
                      <input type="submit" name="button" id="button3" value="Generate Report" />
                    </span>
                  </form></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF">Page <font color="3366ff"> 1 </font>of <font color="3366ff"> 110 </font>: <font color="#000000"><b>1</b></font> <a href='global_ma_Seasonal_approve.asp?num=50&amp;orderstatus=Open'>2</a> <a href='global_ma_Seasonal_approve.asp?num=100&amp;orderstatus=Open'>3</a> <a href='global_ma_Seasonal_approve.asp?num=150&amp;orderstatus=Open'>4</a> <a href='global_ma_Seasonal_approve.asp?num=200&amp;orderstatus=Open'>5</a> <a href='global_ma_Seasonal_approve.asp?num=250&amp;orderstatus=Open'>6</a> <a href='global_ma_Seasonal_approve.asp?num=300&amp;orderstatus=Open'>7</a> <a href='global_ma_Seasonal_approve.asp?num=350&amp;orderstatus=Open'>8</a> <a href='global_ma_Seasonal_approve.asp?num=400&amp;orderstatus=Open'>9</a> <a href='global_ma_Seasonal_approve.asp?num=450&amp;orderstatus=Open'>10</a> <a href='global_ma_Seasonal_approve.asp?num=500&amp;orderstatus=Open'>11</a> <a href='global_ma_Seasonal_approve.asp?num=550&amp;orderstatus=Open'>12</a> <a href='global_ma_Seasonal_approve.asp?num=600&amp;orderstatus=Open'>13</a> <a href='global_ma_Seasonal_approve.asp?num=650&amp;orderstatus=Open'>14</a> <a href='global_ma_Seasonal_approve.asp?num=700&amp;orderstatus=Open'>15</a> <a href='global_ma_Seasonal_approve.asp?num=750&amp;orderstatus=Open'>16</a></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="6%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="20%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Product Code.</span></strong></font></td>
                      <td width="29%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Product Name</span></strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> HQ1</span></strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> HQ2<br />
                      </span></strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>BC</strong></font></td>
                      <td width="6" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>WC</span></strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>NC</span></strong></font></td>
                      <td width="6" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>AC </span></strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>PC</span></strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>JC</span></strong></font></td>
                      <td width="8" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td align="left" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td align="left" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td height="40" align="center"> 1 </td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF">Part0133 </font></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"> Spare-Part 01</td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="rm_rpt_inventory_shockin_detail.asp" target="_blank">12</a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><a href="rm_rpt_inventory_shockin_detail.asp">-12</a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong> 12</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong>-12</strong></td>
                      <!--Open-->
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> 32 </strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>-12</strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>32 </strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>-12</strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>32 </strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>-12</strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>32 </strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>-12</strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>32 </strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>-12</strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>32 </strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>-12</strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong> 150 </strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>-150</strong></td>
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="5" colspan="21" align="right" bgcolor="#FFFFFF"><hr /></td>
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="3" align="right" bgcolor="#FFFFFF"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> 21</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>-12</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong> 21</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong>-12</strong></td>
                      <!--Open-->
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>21</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>-12</strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>32 </strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>-12</strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>32 </strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>-12</strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>32 </strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>-12</strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>32 </strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>-12</strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>32 </strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong>-12</strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong> 200 </strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong>-200</strong></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF">Page <font color="3366ff"> 1 </font>of <font color="3366ff"> 110 </font>: <font color="#000000"><b>1</b></font> <a href='global_ma_Seasonal_approve.asp?num=50&amp;orderstatus=Open'>2</a> <a href='global_ma_Seasonal_approve.asp?num=100&amp;orderstatus=Open'>3</a> <a href='global_ma_Seasonal_approve.asp?num=150&amp;orderstatus=Open'>4</a> <a href='global_ma_Seasonal_approve.asp?num=200&amp;orderstatus=Open'>5</a> <a href='global_ma_Seasonal_approve.asp?num=250&amp;orderstatus=Open'>6</a> <a href='global_ma_Seasonal_approve.asp?num=300&amp;orderstatus=Open'>7</a> <a href='global_ma_Seasonal_approve.asp?num=350&amp;orderstatus=Open'>8</a> <a href='global_ma_Seasonal_approve.asp?num=400&amp;orderstatus=Open'>9</a> <a href='global_ma_Seasonal_approve.asp?num=450&amp;orderstatus=Open'>10</a> <a href='global_ma_Seasonal_approve.asp?num=500&amp;orderstatus=Open'>11</a> <a href='global_ma_Seasonal_approve.asp?num=550&amp;orderstatus=Open'>12</a> <a href='global_ma_Seasonal_approve.asp?num=600&amp;orderstatus=Open'>13</a> <a href='global_ma_Seasonal_approve.asp?num=650&amp;orderstatus=Open'>14</a> <a href='global_ma_Seasonal_approve.asp?num=700&amp;orderstatus=Open'>15</a> <a href='global_ma_Seasonal_approve.asp?num=750&amp;orderstatus=Open'>16</a></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
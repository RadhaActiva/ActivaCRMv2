<!-- #include file="header.asp" -->
<%
set rs = server.CreateObject("adodb.recordset")  
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
                      <td class="titlegrey1"> Stock-Out Report                          
                        <label for="select"></label></td>
                      <td width="196" align="center" class="titlegrey1"><img src="images/excel.jpg" width="57" height="21" /></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
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
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> HQ<br />
                      </span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>BC</strong></font></td>
                      <td width="6%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>WC</span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>NC</span></strong></font></td>
                      <td width="6%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>AC </span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>PC</span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>JC</span></strong></font></td>
                      <td width="8%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td height="40" align="center"> 1 </td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF">Part0133 </font></strong></td>
                      <td align="left" nowrap="nowrap"> Spare-Part 01</td>
                      <td align="center" nowrap="nowrap"><strong> 12</strong></td>
                      <!--Open-->
                      <td align="center" nowrap="nowrap"><strong> 32 </strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong> 150 </strong></td>
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="40" align="center"> 2 </td>
                      <td align="left" nowrap="nowrap"><strong><font color="#0000FF">Part0136</font></strong></td>
                      <td align="left">Spare-Part 02</td>
                      <td align="center" nowrap="nowrap"><strong> 21</strong></td>
                      <!--Open-->
                      <td align="center" nowrap="nowrap"><strong>21</strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong>32 </strong></td>
                      <td align="center"><strong> 200 </strong></td>
                    </tr>
                    <tr>
                      <td colspan="10" align="right"><strong> Total</strong></td>
                      <td align="center"><strong> 350</strong></td>
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
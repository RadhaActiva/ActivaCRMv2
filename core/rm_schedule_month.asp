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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Schedule</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="global_ma_Seasonal_approve.asp?type=searchdata">
                    <table border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td><label for="select"></label>
                          <select name="select" id="select">
                            <option>Technician Code</option>
                            <option>Technician Name</option>
                          </select>
                          <input type="text" name="searchvalue" id="searchvalue" value="" />
                          <input type="submit" name="button" id="button3" value="Submit" />
                          <input type="hidden" name="OrderStatus" value="Open" />
                          (<a href="rm_schedule_month.asp"><strong>View by Month</strong></a> | <strong><a href="rm_schedule.asp">View by Day</a></strong>)</td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
                    <tr>
                      <td class="style21"><strong>Technician Schedule </strong></td>
                    </tr>
                    <tr>
                      <td valign="top"><table width="99%" border="0" cellpadding="1" cellspacing="0">
                        <tr>
                          <td valign="top" bgcolor="#EBEBEB"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="15"><a href="/webapp/www/administration_news.php?yr=2015&amp;mth=10"><img src="images/calendarleft.jpg" width="32" height="20" border="0" /></a></td>
                              <td align="center"><font face="Verdana"><b>Nov 2015
                                <input type="button" name="Button" value="Print Friendly" onclick="javascript:popup('administration_news-popup.php?yr=2015&amp;mth=','Calendar','scrollbars=yes,resizable=yes,width='+screenwidth+',height='+screenHeight)" />
                              </b></font></td>
                              <td width="15"><a href="/webapp/www/administration_news.php?yr=2015&amp;mth=12"><img src="images/calendarright.jpg" width="32" height="20" border="0" /></a></td>
                            </tr>
                          </table>
                            <table class="calendar" width="100%" cellspacing="1" cellpadding="3">
                              <tr>
                                <th abbr="Sunday" bgcolor="#909398" width="14%">Sun</th>
                                <th abbr="Monday" bgcolor="#909398" width="14%">Mon</th>
                                <th abbr="Tuesday" bgcolor="#909398" width="14%">Tue</th>
                                <th abbr="Wednesday" bgcolor="#909398" width="14%">Wed</th>
                                <th abbr="Thursday" bgcolor="#909398" width="14%">Thu</th>
                                <th abbr="Friday" bgcolor="#909398" width="14%">Fri</th>
                                <th abbr="Saturday" bgcolor="#909398" width="14%">Sat</th>
                              </tr>
                              <tr valign="top">
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=1','bible','scrollbars=yes,resizable=no,width=550,height=580');">1</a></b><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2900','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#D014FF;font-size:11px;'><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2876','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 9 am, Cheras</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=2','bible','scrollbars=yes,resizable=no,width=550,height=580');">2</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2876','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 9 am, Puchong</b></a><a href="javascript:popup('appointment_month_view.php?id=2772','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=3','bible','scrollbars=yes,resizable=no,width=550,height=580');">3</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2777','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#D014FF;font-size:11px;'><b>ED only start at 1:30</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#F6AB00;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=4','bible','scrollbars=yes,resizable=no,width=550,height=580');">4</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2781','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=5','bible','scrollbars=yes,resizable=no,width=550,height=580');">5</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2781','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=6','bible','scrollbars=yes,resizable=no,width=550,height=580');">6</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2781','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=7','bible','scrollbars=yes,resizable=no,width=550,height=580');">7</a></b><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2669','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#FF6600;font-size:11px;'><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2886','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Sylvia Chiew start at 12:30pm</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2877','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 1:30pm</b><br />
                                  </a></td>
                              </tr>
                              <tr valign="top" >
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=8','bible','scrollbars=yes,resizable=no,width=550,height=580');">8</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2672','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#FF6600;font-size:11px;'><b>K.L. Prof Selangor Start at 10:30am</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2887','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Sylvia Chiew start at 12:30pm</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2878','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 9 am</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=9','bible','scrollbars=yes,resizable=no,width=550,height=580');">9</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2876','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 9 am, KLCC</b></a><a href="javascript:popup('appointment_month_view.php?id=2773','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=10','bible','scrollbars=yes,resizable=no,width=550,height=580');">10</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2711','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><b>Deepavali Public Holiday</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2778','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#D014FF;font-size:11px;'><b>ED only start at 1:30</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=11','bible','scrollbars=yes,resizable=no,width=550,height=580');">11</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2786','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2879','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 10:30pm</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=12','bible','scrollbars=yes,resizable=no,width=550,height=580');">12</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2786','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2879','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 10:30pm</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=13','bible','scrollbars=yes,resizable=no,width=550,height=580');">13</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2861','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#D014FF;font-size:11px;'><b>ED Only</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2897','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><b>Rayson Annual Leave</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=14','bible','scrollbars=yes,resizable=no,width=550,height=580');">14</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2787','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#D014FF;font-size:11px;'><b>ED Only</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2897','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><b>Rayson Annual Leave</b><br />
                                  </a></td>
                              </tr>
                              <tr valign="top" >
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=15','bible','scrollbars=yes,resizable=no,width=550,height=580');">15</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2854','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><b>Substitute 16/9/15 Malaysia Selangor</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2853','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><b>Off Selangor</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=16','bible','scrollbars=yes,resizable=no,width=550,height=580');">16</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2876','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 9 am, P.J</b></a><a href="javascript:popup('appointment_month_view.php?id=2774','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=17','bible','scrollbars=yes,resizable=no,width=550,height=580');">17</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2779','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#D014FF;font-size:11px;'><b>ED only start at 1:30</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2898','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><b>Rayson Annual Leave</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=18','bible','scrollbars=yes,resizable=no,width=550,height=580');">18</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2783','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=19','bible','scrollbars=yes,resizable=no,width=550,height=580');">19</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2783','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=20','bible','scrollbars=yes,resizable=no,width=550,height=580');">20</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2783','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=21','bible','scrollbars=yes,resizable=no,width=550,height=580');">21</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2670','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#FF6600;font-size:11px;'><b>K.L. Prof Selangor Start at 12:30pm</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><b>School Holiday</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2889','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Sylvia Chiew start at 12:30pm</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2880','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 1:30pm</b><br />
                                  </a></td>
                              </tr>
                              <tr valign="top" >
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=22','bible','scrollbars=yes,resizable=no,width=550,height=580');">22</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2671','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#FF6600;font-size:11px;'><b>K.L. Prof Selangor Start at 10:30am</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><b>School Holiday</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2888','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Sylvia Chiew start at 12:30pm</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2881','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 9 am</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=23','bible','scrollbars=yes,resizable=no,width=550,height=580');">23</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><b>School Holiday</b><br />
                                    </a><br />
                                    <a href="javascript:popup('appointment_month_view.php?id=2876','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 9 am, Klang</b></a><a href="javascript:popup('appointment_month_view.php?id=2775','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=24','bible','scrollbars=yes,resizable=no,width=550,height=580');">24</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><b>School Holiday</b><br />
                                    </a><br />
                                    <a href="javascript:popup('appointment_month_view.php?id=2876','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 9 am, Cheras</b></a><a href="javascript:popup('appointment_month_view.php?id=2780','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#D014FF;font-size:11px;'><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=25','bible','scrollbars=yes,resizable=no,width=550,height=580');">25</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><b>School Holiday</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2784','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2893','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 10:30pm</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=26','bible','scrollbars=yes,resizable=no,width=550,height=580');">26</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><b>School Holiday</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2784','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2893','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 10:30pm</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=27','bible','scrollbars=yes,resizable=no,width=550,height=580');">27</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><b>School Holiday</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2784','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#0000FF;font-size:11px;'><b>Techician01 Selangor (Service - S)</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=28','bible','scrollbars=yes,resizable=no,width=550,height=580');">28</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><b>School Holiday</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2856','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#FF6600;font-size:11px;'><b>Prof Service Talk at Penang</b><br />
                                  </a></td>
                              </tr>
                              <tr valign="top" >
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=29','bible','scrollbars=yes,resizable=no,width=550,height=580');">29</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><b>School Holiday</b><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2855','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#FF6600;font-size:11px;'><b>Prof Service Talk at Penang</b><br />
                                  </a></td>
                                <td bgcolor='#FFFFFF' style='background-color:#FFFFFF;color:#999;' height='100'><b><a href="javascript:popup('appointment_month_view.php?yr=2015&amp;mth=11&amp;day=30','bible','scrollbars=yes,resizable=no,width=550,height=580');">30</a></b><br />
                                  <br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2876','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#994596;font-size:11px;'><b>Technician 02 start at 9 am, Cheras</b></a><a href="javascript:popup('appointment_month_view.php?id=2712','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#666666;font-size:11px;'><br />
                                    </a><br />
                                  <a href="javascript:popup('appointment_month_view.php?id=2776','news','scrollbars=yes,resizable=no,width=550,height=580');" style='color:#F00E0E;font-size:11px;'><b>Off Selangor</b><br />
                                  </a></td>
                                <td colspan="5" style="background-color:#E6E6E6">&nbsp;</td>
                              </tr>
                            </table></td>
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
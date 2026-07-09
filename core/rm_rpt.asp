<!-- #include file="header.asp" -->
<%
set rs = server.CreateObject("adodb.recordset")  

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View</font> 
                        Reports</div></td>
                      </tr>
                    </table></td>
                </tr>
               <tr>
                <th scope="row">&nbsp;</th>
             </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="98%" border="0" cellspacing="0" cellpadding="3">
                    <tr>
                      <td width="46%" align="left" bgcolor="#CCCCCC" scope="row"><strong>Failure Analysis Report</strong></td>
                      <td align="left" bgcolor="#CCCCCC"><strong>Finance</strong></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" scope="row"><table width="98%" border="1" cellspacing="0" cellpadding="3">
                        <tr>
                          <td align="center" scope="row"><strong>1.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_farmonth_year.asp?type=reset" target="_blank">By Month - Faulty Code (One Year)</a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>2.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_farmonth_spareparts.asp?type=reset" target="_blank">By Month - Spare Parts (One Year)</a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>3.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_fardaterange.asp?type=reset" target="_blank">By Date Range</a></strong></td>
                        </tr>    
                        <tr>
                          <td align="left" bgcolor="#CCCCCC" colspan="3" scope="row"><strong>Report By Most Problem</strong></td>
                          </tr>
                        <tr>
                         <tr>
                          <td align="center" scope="row"><strong>1.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_10mostprobUnder.asp?type=reset" target="_blank">Most Problem Under Warranty (By Model)</a></strong></td>
                          </tr>
                        <tr>
                          <td align="center" scope="row"><strong>2.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_10mostprobOver.asp?type=reset" target="_blank">Most Problem Over Warranty (By Model)</a></strong></td>
                          </tr>
                        <tr>
                          <td align="center" scope="row"><strong>3.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_10mostprobPartsUnder.asp" target="_blank">Most Problem Under Warranty (By Parts)</a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>4.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_10mostprobPartsOver.asp" target="_blank">Most Problem Over Warranty (By Parts)</a></strong></td>
                        </tr>
                         <tr>
                          <td align="left" bgcolor="#CCCCCC" colspan="3" scope="row"><strong>Technicians Performance</strong></td>
                          </tr>
                        <tr>
                          <td align="center" scope="row"><strong>1.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_monthtech.asp?type=reset" target="_blank">Monthly Tech Report</a></strong></td>
                        </tr>                       
                        <tr>
                          <td align="center" scope="row"><strong>2.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_tech_servicekpi.asp?type=reset" target="_blank">Service KPI Summary</a></strong></td>
                        </tr>
                        <tr>
                          <td align="left" bgcolor="#CCCCCC" colspan="3" scope="row"><strong>Technicians Claims</strong></td>
                         </tr>
                           <tr>
						 <td align="center" scope="row"><strong>1.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_weekly_job_submission.asp?type=reset" target="_blank">Weekly Service Job Submission</a></strong></td>
                          </tr>
                          <tr>
                          <td align="center" scope="row"><strong>2.</strong></td>
                          <td scope="row"><strong><a href="rm_technician_Claim_Adjustment.asp?type=reset" target="_blank">Monthly Claims Adjustment</a></strong></td>
                          </tr>
                        <tr>
                          <td align="center" scope="row"><strong>3.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_IncentiveTechnician_new.asp" target="_blank">Generate Incentive (Over-Warranty)</a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>4.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_tech_monthcommisionIHT.asp?type=reset" target="_blank">Generate Claims - (In House-Tech)</a></strong></td>
                          </tr>
                        <tr>
                          <td align="center" scope="row"><strong>5.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_tech_monthcommisionTPC.asp?type=reset" target="_blank">Generate Claims - (3rd Party Contractors)</a></strong></td>
                        </tr>                         
                          <tr>
                          <td align="center" scope="row"><strong>6.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_tech_monthcommisionIC.asp?type=reset" target="_blank">Generate Claims - (Independent Contractors)</a></strong></td>
                        </tr>                         
                          <tr>
                          <td align="center" scope="row"><strong>7.</strong></td>
                          <td scope="row"><strong><a href="rmtech_claims_manual.asp?type=reset" target="_blank">Manual Claims Entry - (IHT)</a></strong></td>
                        </tr>  
                      </table></td>

                      <td align="left" valign="top">
                          <table width="98%" border="1" cellspacing="0" cellpadding="3">
                              <tr>
                          <td width="11%" align="center" scope="row"><strong>1.</strong></td>
                          <td width="89%" scope="row"><a href="rm_rpt_pnL.asp?type=reset" target="_blank"><strong>Monthly Inventory Assessment Report</strong></a></td>
                        </tr>
                          <tr>
                          <td align="center" scope="row"><strong>2.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_salesanalysis_summary.asp?type=reset" target="_blank">Sales Invoice Summary (Cross Month) </a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>3.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_aging_statement.asp?type=reset" target="_blank">Debtor Aging Report </a></strong></td>
                        </tr>
                        <td align="center" scope="row"><strong>4.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_cn_summary.asp?type=reset" target="_blank">CN Summary Report</a></strong></td>
                        </tr>
                          <tr>
                          <td align="left" bgcolor="#CCCCCC" colspan="3" scope="row"><strong>Inventory</strong></td>
                         </tr>
                        <tr>
                          <td align="center" scope="row"><strong>1.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_inventory_warehouse_location.asp?type=reset" target="_blank">Inventory By Store Location</a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>2.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_inventory_productgroup.asp?type=reset" target="_blank">Inventory By Product Group</a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>3.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_stockcard_ageing.asp?type=reset" target="_blank">Stock Ageing</a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>4.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_stockcard_warehouse.asp?type=reset" target="_blank">Summary By Store Location</a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>5.</strong></td>
                          <td scope="row"><strong><a href="rm_rpt_stockcard_ledger.asp?type=reset" target="_blank">Stock Ledger by Stock Code</a></strong></td>
                        </tr>
                         <tr>
                          <td align="center" scope="row"><strong>6.</strong></td>
                          <td scope="row"><a href="rm_rpt_stock_movement.asp?type=reset" target="_blank"><strong>Stock Movement by Stock Code</strong></a></td>
                        </tr>
                          <tr>
                          <td align="center" scope="row"><strong>7.</strong></td>
                          <td scope="row"><a href="rm_rpt_under_wrty_parts.asp?type=reset" target="_blank"><strong>Under Warranty Spare Parts</strong></a></td>
                        </tr>
                           <tr>
                          <td align="center" scope="row"><strong>8.</strong></td>
                          <td scope="row"><a href="rm_rpt_stockcard_ledger_stockrange.asp?type=reset" target="_blank"><strong>Stock Ledger by Stock Range</strong></a></td>
                        </tr>
                           <tr>
                          <td align="center" scope="row"><strong>9.</strong></td>
                          <td scope="row"><a href="rm_por_rpt.asp?type=reset" target="_blank"><strong>POP Report</strong></a></td>
                        </tr>
                            </tr>    
                       <tr>
                          <td align="left" bgcolor="#CCCCCC" colspan="3" scope="row"><strong>Dealer</strong></td>
                         </tr>
                             <tr>
                          <td align="center" scope="row"><strong>10.</strong></td>
                          <td scope="row"><a href="rm_rpt_dealer.asp?type=reset" target="_blank"><strong>Dealer Commission Report</strong></a></td>
                        </tr> 
                   <tr>
                          <td align="center" scope="row"><strong>11.</strong></td>
                          <td scope="row"><a href="rm_rpt_dealer_collection.asp?type=reset" target="_blank"><strong>Dealer Collection Report</strong></a></td>
                        </tr>     
                      </table></td>
                    </tr>
                    <tr>
                      <td align="left" scope="row">&nbsp;</td>
                      <td align="left">&nbsp;</td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td width="89" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                  <td width="892">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
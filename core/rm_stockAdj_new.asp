<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

if request("sj_no") <> "" then	  
sql = "SELECT sj_id, sj_no, sj_date, sj_referenceno, sj_status, sj_fromwarehouse, sj_towarehouse, sj_remark, sj_createddate, sj_createdby, sj_submitteddate, " & _
		"sj_submittedby, sj_approveddate, sj_approveddate, sj_approvedby, sj_cancelleddate, sj_cancelledby, sj_totalqty, sj_totalaAmt, sj_emailsent, sj_emailsentdate " & _
		"FROM tblstockadj WHERE sj_no = '" & request("sj_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sj_id = rs("sj_id") 
			sj_no = rs("sj_no") 
			sj_date = rs("sj_date")
			sj_referenceno = rs("sj_referenceno") 
			sj_status = rs("sj_status")
			sj_fromwarehouse = rs("sj_fromwarehouse") 
			sj_towarehouse = rs("sj_towarehouse")
			sj_remark = rs("sj_remark")
			sj_createddate = rs("sj_createddate") 
			sj_createdby = rs("sj_createdby") 
			sj_submitteddate = rs("sj_submitteddate") 
			sj_submittedby = rs("sj_submittedby") 
			sj_approveddate = rs("sj_approveddate") 
			sj_approvedby = rs("sj_approvedby") 
			sj_cancelleddate = rs("sj_cancelleddate") 
			sj_cancelledby = rs("sj_cancelledby") 
			sj_totalqty = rs("sj_totalqty") 
			sj_totalaAmt = rs("sj_totalaAmt") 
			sj_emailsent = rs("sj_emailsent")  
			sj_emailsentdate = rs("sj_emailsentdate")
		End If
		rs.Close
	  stype = "editStockAdj"	
	  actionname = "Save" 
 else    
	  stype = "addStockAdj"
	  actionname = "Save" 		
	  sj_date = date()	 
	  sj_status = "Open"   	
end if
%> 
<script language="javascript">

function confirmForm(id,orderlinks,otype) 
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formdodetail.action = orderlinks;
	document.formdodetail.submit();
   }
}


function calctotal(unitprice,qty,discountamt,discounttype,subtotal) {

var discount = 0;
var temp = 0;

if (discounttype == "%")
{
	discount = unitprice * (discountamt/100);
	temp = (unitprice-discount) * qty;
	subtotal.value = temp.toFixed(2);
	}
else
{
	discount = discountamt;
	temp = (unitprice-discount) * qty;
	subtotal.value = temp.toFixed(2);
	}
}
// -->
</script>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Create/Edit </font>Stock Adjustment (ADJ)</div></td>
                        <td align="right" class="titleblue1"><a href="rm_stockAdj_new_print.asp?sj_no=<%=sj_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/></a></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" id='ticketTbl'>
                    <thead>
                      <tr class="head_row">
                        <td colspan="2" nowrap="nowrap" bgcolor="#DADADA"><strong>Stock-Adjustment Information </strong></td>
                      </tr>
                    </thead>
                    
                    <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                      <tr class="head_row">
                        <td colspan="2" valign="top"><table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#EBEBEB">
                          <tr>
                            <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Stock-ADJ No</strong><br />
                             <strong><font size="1">(System Generate) </font></strong></font></td>
                            <td nowrap="nowrap"><strong><%=sj_no%><input name="sj_no" type="hidden" id="sj_no" value="<%=sj_no%>" />
                            </strong></td>
                            <td height="22" nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Stock-ADJ date</strong></font></strong></td>
                            <td><strong>
                              <input name="sj_date" type="text" id="sj_date" value="<%=chkdate(sj_date)%>" size="15" maxlength="20" />
                              <font color="#000000"><strong><a href="javascript:void(null)" onclick="window.dateField = document.formorder.do_purchase_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong></td>
                            <td bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong> Store</strong></font></strong></td>
                            <td><input name="sj_fromwarehouse" type="text" id="stkIn_warehouse" value="<%=sj_fromwarehouse%>" size="25" maxlength="50" onfocus="this.blur();" />
                              [<a href="javascript:popup('rm_stockin_new_warehouse.asp?searchitem=wh_code&amp;searchvalue=<%=cusj_code%>&formname=formorder&fieldname=sj_fromwarehouse','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>]</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Reference No</strong></font></strong></td>
                            <td nowrap="nowrap"><input name="sj_referenceno" type="text" id="sj_referenceno" value="<%=sj_referenceno%>" size="20" maxlength="100" /></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong> Status</strong></font></strong></td>
                            <td align="left" nowrap="nowrap"><%=sj_status%></td>
                            <td align="left" nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                            <td align="left" nowrap="nowrap">&nbsp;</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Created by</strong></font></strong></td>
                            <td nowrap="nowrap"><%=sj_createdby%> @ <%=chkdatetime(sj_createddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Remark</strong></font></strong></td>
                            <td colspan="3"><input name="sj_remark" type="text" id="supp_code" value="<%=sj_remark%>" size="50" maxlength="150" /></td>
                            </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Submitted By</strong></font></strong></td>
                            <td nowrap="nowrap"><%=sj_submittedby%> @ <%=chkdatetime(sj_submitteddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Cancelled by</strong></font></strong></td>
                            <td><%=sj_cancelledby%> @ <%=chkdatetime(sj_cancelleddate)%></td>
                            <td colspan="2" align="right">&nbsp;</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Approved By</strong></font></strong></td>
                            <td nowrap="nowrap"><%=sj_approvedby%> @ <%=chkdatetime(sj_approveddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td colspan="2" align="right"><%if sj_status<>"Approved" and sj_status<>"Cancel" then %>                              <input type="submit" name="button" id="button" value="<%=actionname%>" />
                              <%end if%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr class="head_row">
                        <td colspan="2" valign="top">&nbsp;</td>
                      </tr>
                    </form>
                    

 
                    <%if sj_no <> "" then %>
                    
<%

if request("sjd_id") <> "" then
		sql = "SELECT sjd_id, sjd_sj_no, sjd_itm_code, sjd_itm_desc, sjd_unitcost, sjd_current_qty, sjd_adjust_qty, sjd_diff_qty , sjd_subtotal, sjd_referid " & _
	          "FROM tblstockadj_detail where sjd_id = '" & request("sjd_id") & "'"	
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   sjd_id = rs("sjd_id")
		   sjd_sj_no = rs("sjd_sj_no")
		   sjd_itm_code = rs("sjd_itm_code")
		   sjd_itm_desc = rs("sjd_itm_desc")
		   sjd_unitcost = rs("sjd_unitcost")
		   sjd_current_qty = rs("sjd_current_qty")
		   sjd_adjust_qty = rs("sjd_adjust_qty")
		   sjd_diff_qty = rs("sjd_diff_qty")
		   sjd_subtotal = rs("sjd_subtotal")
		   sjd_referid = rs("sjd_referid")
        end if
		rs.close
		sbutton = "Update"
		stype="editStockAdjDetail"	
else
		sbutton = "Add"
		stype="addStockAdjDetail"
		sjd_current_qty = "0"	
		sjd_adjust_qty = "0"	
		sjd_new_qty = "0"	
		sjd_unitcost = "0.00"	
		sjd_subtotal = "0.00"
end if

%>
                    <tr class="head_row">
                      <td colspan="2"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E8E8E8">
                      <tr class="head_row">
                          <td width="3%" height="24" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>No</strong></font></strong></td>
                          <td width="27%" align="left" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Item Code</strong></font></strong></td>
                          <td width="17%" align="left" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Item Name / Description</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong> Current Qty</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong> ADJ Qty</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Diff Qty</strong></font></strong></td>
                          <td width="20%" align="right" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Action</strong></font></strong></td>
                        </tr>
                        
                       <%if sj_status="Open" then %>
                       <form name="formdodetail" id="formdodetail" method="post" action="rm_jobsheet.asp#spareparts" >
                        <tr class="head_row">
                          <td height="24" colspan="2" bgcolor="#666666"><input name="sjd_itm_code" type="text" id="sjd_itm_code" value="<%=sjd_itm_code%>" size="25" maxlength="50" onfocus="this.blur();" />
                            <a href="javascript:popup('rm_stockAdj_new_item.asp?searchitem=md_code&searchvalue=&amp;sj_fromwarehouse='+document.formorder.sj_fromwarehouse.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">[Select]</a></td>
                          <td align="left" bgcolor="#666666"><input name="sjd_itm_desc" type="text" id="sjd_itm_desc" value="<%=sjd_itm_desc%>" size="30" maxlength="100" />
                            <input type="hidden" name="sj_no" id="sj_no" value="<%=sj_no%>" />
                            <input type="hidden" name="sjd_id" id="sjd_id" value="<%=sjd_id%>" />
                          </td>
                          <td align="center" bgcolor="#666666"><input name="sjd_current_qty" type="text" id="sjd_current_qty" style="text-align: right;" value="<%=sjd_current_qty%>" size="4" maxlength="10" /></td>
                          <td align="center" bgcolor="#666666"><input name="sjd_adjust_qty" type="text" id="sjd_adjust_qty" style="text-align: right;" value="<%=sjd_adjust_qty%>" size="4" maxlength="10" /></td>
                          <td align="center" bgcolor="#666666">&nbsp;</td>
                          <td align="center" bgcolor="#666666"><input type="button" name="button2" id="button2" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=request("dod_id")%>','action.asp?type=<%=stype%>','<%=dod_subtotal%>');" /></td>
                        </tr>
                        </form>
                         <%end if%> 
                         
 <%				i = 1
				sql1 = "SELECT sjd_id, sjd_sj_no, sjd_itm_code, sjd_itm_desc, sjd_unitcost, sjd_current_qty, sjd_adjust_qty, sjd_diff_qty, sjd_subtotal, sjd_referid " & _
	                   "FROM tblstockadj_detail where sjd_sj_no = '" & sj_no & "' order by sjd_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
                         
                        <tr>
                          <td align="center"><%=i%>.</td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("sjd_itm_code")%></td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("sjd_itm_desc")%></td>
                          <td align="center" bgcolor="#FFFFFF"><%=rs1("sjd_current_qty")%></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><%=rs1("sjd_adjust_qty")%></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><%=rs1("sjd_diff_qty")%></td>
                          <td align="center" nowrap="nowrap" bgcolor="#FFFFFF" class='tktTotals'><%if sj_status="Open" then %>
                            <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_stockAdj_new.asp?sjd_id=<%=rs1("sjd_id")%>&amp;sj_no=<%=rs1("sjd_sj_no")%>#spareparts'" />
                            <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmAction('<%=rs1("sjd_itm_code")%>','action.asp?type=delStockAdjDetail&sjd_id=<%=rs1("sjd_id")%>&sj_no=<%=rs1("sjd_sj_no")%>')" />
                            <%end if%></td>
                        </tr>
 <%	
				sjd_current_qty = sjd_current_qty + rs1("sjd_current_qty")
				sjd_adjust_qty = sjd_adjust_qty + rs1("sjd_adjust_qty")
				sjd_diff_qty = sjd_diff_qty + rs1("sjd_diff_qty")
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%> 
                        
                        <tr valign="top">
                          <td colspan="3" align="right" bgcolor="#FFFFFF"><strong> Total Qty:</strong></td>
                          <td align="center" bgcolor="#FFFFFF"><strong><%=sjd_current_qty%></strong></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><strong><%=sjd_adjust_qty%></strong></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'>&nbsp;</td>
                          <td align="center" nowrap="nowrap" bgcolor="#FFFFFF" class='tktTotals'>&nbsp;</td>
                        </tr>
                        <tr>
                          <td colspan="7" valign="top"><table width="100%">
                            <tr>
                              <td width="50%" valign="top" bgcolor="#FFFFFF"><%if sj_status<>"Approved" and sj_status<>"Cancel" then %>
                                <input type="button" name="DoneInvoice2" id="DoneInvoice2" value="Cancel Stock-Adj" onclick="javascript:confirmAction('<%=sj_no%>','action.asp?type=CancelStockAdj&amp;sj_no=<%=sj_no%>')" />
                                <%end if%>
                                <br />
                                <br /></td>
                              <td align="right" valign="top" bgcolor="#FFFFFF"><%if sj_status="Open" then %>
                                <input type="button" name="SubmitJob2" id="SubmitJob2" value="Submit Stock-Adj " onclick="javascript:confirmAction('<%=sj_no%>','action.asp?type=SubmitStockAdj&amp;sj_no=<%=sj_no%>')" />
                                <%end if%>
                                <br />
                                <%if sj_status="Submitted" then %>

                                  <%if Request.Cookies("GAPS")("approve_stk") = "Y" then %>
                                        <input type="button" name="DoneInvoice" id="DoneInvoice" value="Approve Stock-Adj" onclick="javascript:confirmAction('<%=sj_no%>','action.asp?type=ApproveStockAdj&amp;sj_no=<%=sj_no%>')" />
                                    <%else %>
                                        <input type="button" name="DoneInvoice" id="DoneInvoice3" value="Approve - Not Authorized" disabled" />
                                  <%end if%>
                                  <%end if%>
                                <br />
                                <br />
                                <br />
                                <br /></td>
                              </tr>
                            </table></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td width="35%"></thead></td>
                    </tr>
                  </table></td>
                </tr>
                <%end if%>
              <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
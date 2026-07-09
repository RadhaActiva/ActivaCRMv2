<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

if request("so_no") <> "" then	  
sql = "SELECT so_id, so_no, so_date, so_referenceno, so_status, so_fromwarehouse, so_towarehouse, so_remark, so_createddate, so_createdby, so_submitteddate, " & _
		"so_submittedby, so_approveddate, so_approveddate, so_approvedby, so_cancelleddate, so_cancelledby, so_totalqty, so_totalaAmt, so_emailsent, so_emailsentdate " & _
		"FROM tblstockOut WHERE so_no = '" & request("so_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			so_id = rs("so_id") 
			so_no = rs("so_no") 
			so_date = rs("so_date")
			so_referenceno = rs("so_referenceno") 
			so_status = rs("so_status")
			so_fromwarehouse = rs("so_fromwarehouse") 
			so_towarehouse = rs("so_towarehouse")
			so_remark = rs("so_remark")
			so_createddate = rs("so_createddate") 
			so_createdby = rs("so_createdby") 
			so_submitteddate = rs("so_submitteddate") 
			so_submittedby = rs("so_submittedby") 
			so_approveddate = rs("so_approveddate") 
			so_approvedby = rs("so_approvedby") 
			so_cancelleddate = rs("so_cancelleddate") 
			so_cancelledby = rs("so_cancelledby") 
			so_totalqty = rs("so_totalqty") 
			so_totalaAmt = rs("so_totalaAmt") 
			so_emailsent = rs("so_emailsent")  
			so_emailsentdate = rs("so_emailsentdate")
		End If
		rs.Close
	  stype = "editStockOut"	
	  actionname = "Save" 
 else    
	  stype = "addStockOut"
	  actionname = "Save" 		
	  so_date = date()	 
	  so_status = "Open"   	
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


function calctotal(unitprice,qty,subtotal) {

    var temp = 0;
	temp = unitprice*qty;
	subtotal.value = temp.toFixed(2);
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Create/Edit </font>Stock Out</div></td>
                        <td align="right" class="titleblue1"><a href="rm_stockOut_new_print.asp?so_no=<%=so_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/></a></td>
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
                        <td colspan="2" nowrap="nowrap" bgcolor="#DADADA"><strong>Stock-Out Information </strong></td>
                      </tr>
                    </thead>
                    
                    <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                      <tr class="head_row">
                        <td colspan="2" valign="top"><table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#EBEBEB">
                          <tr>
                            <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Stock-Out No</strong><br />
                             <strong><font size="1">(System Generate) </font></strong></font></td>
                            <td nowrap="nowrap"><strong><%=so_no%><input name="so_no" type="hidden" id="so_no" value="<%=so_no%>" />
                            </strong></td>
                            <td height="22" nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Stock-Out date</strong></font></strong></td>
                            <td><strong>
                              <input name="so_date" type="text" id="so_date" value="<%=chkdate(so_date)%>" size="15" maxlength="20" />
                              <font color="#000000"><strong><a href="javascript:void(null)" onclick="window.dateField = document.formorder.do_purchase_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong></td>
                            <td bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong> Store</strong></font></strong></td>
                            <td><input name="so_fromwarehouse" type="text" id="stkIn_warehouse" value="<%=so_fromwarehouse%>" size="25" maxlength="50" />
                              [<a href="javascript:popup('rm_stockin_new_warehouse.asp?searchitem=wh_code&amp;searchvalue=<%=cuso_code%>&formname=formorder&fieldname=so_fromwarehouse','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>]</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Reference No</strong></font></strong></td>
                            <td nowrap="nowrap"><input name="so_referenceno" type="text" id="so_referenceno" value="<%=so_referenceno%>" size="20" maxlength="100" /></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong> Status</strong></font></strong></td>
                            <td align="left" nowrap="nowrap"><%=so_status%></td>
                            <td align="left" nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                            <td align="left" nowrap="nowrap">&nbsp;</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Created by</strong></font></strong></td>
                            <td nowrap="nowrap"><%=so_createdby%> @ <%=chkdatetime(so_createddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Remark</strong></font></strong></td>
                            <td colspan="3"><input name="so_remark" type="text" id="supp_code" value="<%=so_remark%>" size="50" maxlength="150" /></td>
                            </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Submitted By</strong></font></strong></td>
                            <td nowrap="nowrap"><%=so_submittedby%> @ <%=chkdatetime(so_submitteddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Cancelled by</strong></font></strong></td>
                            <td><%=so_cancelledby%> @ <%=chkdatetime(so_cancelleddate)%></td>
                            <td colspan="2" align="right">&nbsp;</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Approved By</strong></font></strong></td>
                            <td nowrap="nowrap"><%=so_approvedby%> @ <%=chkdatetime(so_approveddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td colspan="2" align="right"><%if so_status<>"Approved" and so_status<>"Cancel" then %>                              <input type="submit" name="button" id="button" value="<%=actionname%>" />
                              <%end if%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr class="head_row">
                        <td colspan="2" valign="top">&nbsp;</td>
                      </tr>
                    </form>
                    

 
                    <%if so_no <> "" then %>
                    
<%

if request("sod_id") <> "" then
		sql = "SELECT sod_id, sod_so_no, sod_itm_code, sod_itm_desc, sod_unitcost, sod_qty, sod_subtotal, sod_referid " & _
	          "FROM tblstockOut_detail where sod_id = '" & request("sod_id") & "'"	
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   sod_id = rs("sod_id")
		   sod_so_no = rs("sod_so_no")
		   sod_itm_code = rs("sod_itm_code")
		   sod_itm_desc = rs("sod_itm_desc")
		   sod_unitcost = rs("sod_unitcost")
		   sod_qty = rs("sod_qty")
		   sod_subtotal = rs("sod_subtotal")
		   sod_referid = rs("sod_referid")
        end if
		rs.close
		sbutton = "Update"
		stype="editStockOutDetail"	
else
		sbutton = "Add"
		stype="addStockOutDetail"
		sod_qty = "1"	
		sod_unitcost = "0.00"	
		sod_subtotal = "0.00"
end if

%>
                    <tr class="head_row">
                      <td colspan="2"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E8E8E8">
                      <tr class="head_row">
                          <td height="24" nowrap="nowrap" bgcolor="#475387 "><strong><font color="#FFFFFF"><strong>No</strong></font></strong></td>
                          <td align="left" nowrap="nowrap" bgcolor="#475387 "><strong><font color="#FFFFFF"><strong>Item Code</strong></font></strong></td>
                          <td align="left" nowrap="nowrap" bgcolor="#475387 "><strong><font color="#FFFFFF"><strong>Item Name / Description</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387 "><strong><font color="#FFFFFF"><strong>Unit Price</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387 "><strong><font color="#FFFFFF"><strong>Stock Out Qty</strong></font></strong></td>
                          <td align="right" nowrap="nowrap" bgcolor="#475387 "><strong><font color="#FFFFFF"><strong>Sub-Total</strong></font></strong></td>
                          <td align="right" nowrap="nowrap" bgcolor="#475387 "><strong><font color="#FFFFFF"><strong>Action</strong></font></strong></td>
                        </tr>
                        
                       <%if so_status="Open" then %>
                       <form name="formdodetail" id="formdodetail" method="post" action="rm_jobsheet.asp#spareparts" >
                        <tr class="head_row">
                          <td height="24" colspan="2" bgcolor="#666666"><input name="sod_itm_code" type="text" id="sod_itm_code" value="<%=sod_itm_code%>" size="25" maxlength="50" />
                            <a href="javascript:popup('rm_stockOut_new_item.asp?searchitem=md_code&searchvalue=&amp;formname=formdodetail&amp;fieldname=sod_itm_code&so_fromwarehouse='+document.formorder.so_fromwarehouse.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">[Select]</a></td>
                          <td align="left" bgcolor="#666666"><input name="sod_itm_desc" type="text" id="sod_itm_desc" value="<%=sod_itm_desc%>" size="30" maxlength="100" />
                            <input type="hidden" name="so_no" id="so_no" value="<%=so_no%>" />
                            <input type="hidden" name="sod_id" id="sod_id" value="<%=sod_id%>" />
                            </td>
                          <td align="center" bgcolor="#666666"><input name="sod_unitcost" type="text" id="sod_unitcost" value="<%=sod_unitcost%>" style="text-align: right;" size="12" maxlength="15" onkeydown="calctotal(document.formdodetail.sod_unitcost.value, document.formdodetail.sod_qty.value, document.formdodetail.sod_subtotal);" onkeyup="calctotal(document.formdodetail.sod_unitcost.value, document.formdodetail.sod_qty.value, document.formdodetail.sod_subtotal);" /></td>
                          <td align="center" bgcolor="#666666"><input name="sod_qty" type="text" id="sod_qty" style="text-align: right;" value="<%=sod_qty%>" size="4" maxlength="10" onkeydown="calctotal(document.formdodetail.sod_unitcost.value, document.formdodetail.sod_qty.value, document.formdodetail.sod_subtotal);" onkeyup="calctotal(document.formdodetail.sod_unitcost.value, document.formdodetail.sod_qty.value, document.formdodetail.sod_subtotal);" /></td>
                          <td align="right" bgcolor="#666666"><input name="sod_subtotal" type="text" id="sod_subtotal" value="<%=sod_subtotal%>" style="text-align: right;" size="12" maxlength="15" onkeydown="calctotal(document.formdodetail.sod_unitcost.value, document.formdodetail.sod_qty.value, document.formdodetail.sod_subtotal);" onkeyup="calctotal(document.formdodetail.sod_unitcost.value, document.formdodetail.sod_qty.value, document.formdodetail.sod_subtotal);" /></td>
                          <td align="right" bgcolor="#666666"><input type="button" name="button2" id="button2" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=request("dod_id")%>','action.asp?type=<%=stype%>','<%=dod_subtotal%>');" /></td>
                        </tr>
                        </form>
                         <%end if%> 
                         
 <%				i = 1
				sql1 = "SELECT sod_id, sod_so_no, sod_itm_code, sod_itm_desc, sod_unitcost, sod_qty, sod_subtotal, sod_referid " & _
	                  "FROM tblstockOut_detail where sod_so_no = '" & so_no & "' order by sod_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
                         
                        <tr valign="top">
                          <td align="center"><%=i%>.</td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("sod_itm_code")%></td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("sod_itm_desc")%></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><%=chknumber2(rs1("sod_unitcost"))%></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><%=rs1("sod_qty")%></td>
                          <td align="right" bgcolor="#FFFFFF" class='tktTotals'><%=chknumber2(rs1("sod_subtotal"))%></td>
                          <td align="right" nowrap="nowrap" bgcolor="#FFFFFF" class='tktTotals'><%if so_status="Open" then %>
                            <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_stockOut_new.asp?sod_id=<%=rs1("sod_id")%>&amp;so_no=<%=rs1("sod_so_no")%>#spareparts'" />
                            <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmAction('<%=rs1("sod_itm_code")%>','action.asp?type=delStockOutDetail&sod_id=<%=rs1("sod_id")%>&so_no=<%=rs1("sod_so_no")%>')" />
                            <%end if%></td>
                        </tr>
 <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%> 
                        
                        <tr valign="top">
                          <td colspan="3" align="right" bgcolor="#FFFFFF"><strong> Total :</strong></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'>&nbsp;</td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><strong><%=so_totalqty%></strong></td>
                          <td align="right" bgcolor="#FFFFFF" class='tktTotals'><strong><%=chknumber2(so_totalaAmt)%></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#FFFFFF" class='tktTotals'>&nbsp;</td>
                        </tr>
                        <tr>
                          <td colspan="7" valign="top"><table width="100%">
                            <tr>
                              <td width="50%" valign="top" bgcolor="#FFFFFF"><%if so_status<>"Approved" and so_status<>"Cancel" then %>
                                <input type="button" name="DoneInvoice2" id="DoneInvoice2" value="Cancel Stock-Out" onclick="javascript:confirmAction('<%=so_no%>','action.asp?type=CancelStockOut&amp;so_no=<%=so_no%>')" />
                                <%end if%>
                                <br />
                                <br /></td>
                              <td align="right" valign="top" bgcolor="#FFFFFF"><%if so_status="Open" then %>
                                <input type="button" name="SubmitJob2" id="SubmitJob2" value="Submit Stock-Out " onclick="javascript:confirmAction('<%=so_no%>','action.asp?type=SubmitStockOut&amp;so_no=<%=so_no%>')" />
                                <%end if%>
                                <br />
                                <%if so_status="Submitted" then %>
                                    <%if Request.Cookies("GAPS")("approve_stk") = "Y" then %>
                                     <input type="button" name="DoneInvoice" id="DoneInvoice" value="Approve Stock-Out" onclick="javascript:confirmAction('<%=sf_no%>','action.asp?type=ApproveStockOut&amp;so_no=<%=so_no%>')" />
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
                      <td width="35%"></head></td>
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
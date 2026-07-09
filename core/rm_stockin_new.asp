<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

if request("st_no") <> "" then	  
sql = "SELECT st_id, st_no, st_date, st_referenceno, st_status, st_fromwarehouse, st_towarehouse, st_remark, st_createddate, st_createdby, st_submitteddate, " & _
		"st_submittedby, st_approveddate, st_approveddate, st_approvedby, st_cancelleddate, st_cancelledby, st_totalqty, st_totalaAmt, st_emailsent, st_emailsentdate " & _
		"FROM tblstockin WHERE st_no = '" & request("st_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			st_id = rs("st_id") 
			st_no = rs("st_no") 
			st_date = rs("st_date")
			st_referenceno = rs("st_referenceno") 
			st_status = rs("st_status")
			st_fromwarehouse = rs("st_fromwarehouse") 
			st_towarehouse = rs("st_towarehouse")
			st_remark = rs("st_remark")
			st_createddate = rs("st_createddate") 
			st_createdby = rs("st_createdby") 
			st_submitteddate = rs("st_submitteddate") 
			st_submittedby = rs("st_submittedby") 
			st_approveddate = rs("st_approveddate") 
			st_approvedby = rs("st_approvedby") 
			st_cancelleddate = rs("st_cancelleddate") 
			st_cancelledby = rs("st_cancelledby") 
			st_totalqty = rs("st_totalqty") 
			st_totalaAmt = rs("st_totalaAmt") 
			st_emailsent = rs("st_emailsent")  
			st_emailsentdate = rs("st_emailsentdate")
		End If
		rs.Close
	  stype = "editStockIn"	
	  actionname = "Save" 
 else    
	  stype = "addStockIn"
	  actionname = "Save" 		
	  st_date = date()	 
	  st_status = "Open"   	
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Create/Edit </font>Stock In</div></td>
                        <td align="right" class="titleblue1"><a href="rm_stockin_new_print.asp?st_no=<%=st_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/></a></td>
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
                        <td colspan="2" nowrap="nowrap" bgcolor="#DADADA"><strong>Stock-In Information </strong></td>
                      </tr>
                    </thead>
                    
                    <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                      <tr class="head_row">
                        <td colspan="2" valign="top"><table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#EBEBEB">
                          <tr>
                            <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Stock-In No</strong><br />
                             <strong><font size="1">(System Generate) </font></strong></font></td>
                            <td nowrap="nowrap"><strong><%=st_no%><input name="st_no" type="hidden" id="st_no" value="<%=st_no%>" />
                            </strong></td>
                            <td height="22" nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Stock-In date</strong></font></strong></td>
                            <td><strong>
                              <input name="st_date" type="text" id="st_date" value="<%=chkdate(st_date)%>" size="15" maxlength="20" />
                              <font color="#000000"><strong><a href="javascript:void(null)" onclick="window.dateField = document.formorder.do_purchase_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong></td>
                            <td bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong> Store</strong></font></strong></td>
                            <td><input name="st_towarehouse" type="text" id="stkIn_warehouse2" value="<%=st_towarehouse%>" size="25" maxlength="50" onfocus="this.blur();" />
[<a href="javascript:popup('rm_stockin_new_warehouse.asp?searchitem=wh_code&amp;searchvalue=<%=cust_code%>&amp;formname=formorder&amp;fieldname=st_towarehouse','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>]</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Reference No</strong></font></strong></td>
                            <td nowrap="nowrap"><input name="st_referenceno" type="text" id="st_referenceno" value="<%=st_referenceno%>" size="20" maxlength="100" /></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong> Status</strong></font></strong></td>
                            <td align="left" nowrap="nowrap"><%=st_status%></td>
                            <td align="left" nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                            <td align="left" nowrap="nowrap">&nbsp;</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Created by</strong></font></strong></td>
                            <td nowrap="nowrap"><%=st_createdby%> @ <%=chkdatetime(st_createddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Remark</strong></font></strong></td>
                            <td colspan="3"><input name="st_remark" type="text" id="supp_code" value="<%=st_remark%>" size="50" maxlength="150" /></td>
                            </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Submitted By</strong></font></strong></td>
                            <td nowrap="nowrap"><%=st_submittedby%> @ <%=chkdatetime(st_submitteddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Cancelled by</strong></font></strong></td>
                            <td colspan="3"><%=st_cancelledby%> @ <%=chkdatetime(st_cancelleddate)%></td>
                            </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Approved By</strong></font></strong></td>
                            <td nowrap="nowrap"><%=st_approvedby%> @ <%=chkdatetime(st_approveddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td colspan="2" align="right"><%if st_status<>"Approved" and st_status<>"Cancel" then %>                              <input type="submit" name="button" id="button" value="<%=actionname%>" />
                              <%end if%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr class="head_row">
                        <td colspan="2" valign="top">&nbsp;</td>
                      </tr>
                    </form>
                    

 
                    <%if st_no <> "" then %>
                    
<%

if request("std_id") <> "" then
		sql = "SELECT std_id, std_st_no, std_itm_code, std_itm_desc, std_unitcost, std_qty, std_subtotal, std_referid " & _
	          "FROM tblstockin_detail where std_id = '" & request("std_id") & "'"	
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   std_id = rs("std_id")
		   std_st_no = rs("std_st_no")
		   std_itm_code = rs("std_itm_code")
		   std_itm_desc = rs("std_itm_desc")
		   std_unitcost = rs("std_unitcost")
		   std_qty = rs("std_qty")
		   std_subtotal = rs("std_subtotal")
		   std_referid = rs("std_referid")
        end if
		rs.close
		sbutton = "Update"
		stype="editStockInDetail"	
else
		sbutton = "Add"
		stype="addStockInDetail"
		std_qty = "1"	
		std_unitcost = "0.00"	
		std_subtotal = "0.00"
end if

%>
                    <tr class="head_row">
                      <td colspan="2"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E8E8E8">
                      <tr class="head_row">
                          <td height="24" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>No</strong></font></strong></td>
                          <td align="left" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Item Code</strong></font></strong></td>
                          <td align="left" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Item Name / Description</strong></font></strong></td>
                          <td align="right" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Unit Cost</strong></font></strong></td>
                          <td align="right" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Stock In Qty</strong></font></strong></td>
                          <td align="right" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Sub-Total</strong></font></strong></td>
                          <td align="right" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Action</strong></font></strong></td>
                        </tr>
                        
                       <%if st_status="Open" then %>
                       <form name="formdodetail" id="formdodetail" method="post" action="rm_jobsheet.asp#spareparts" >
                        <tr class="head_row">
                          <td height="24" colspan="2" bgcolor="#666666"><input name="std_itm_code" type="text" id="std_itm_code" value="<%=std_itm_code%>" size="25" maxlength="50" />
                            <a href="javascript:popup('rm_stockin_new_item.asp?searchitem=md_code&searchvalue=&amp;formname=formdodetail&amp;fieldname=std_itm_code&','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">[Select]</a></td>
                          <td align="left" bgcolor="#666666"><input name="std_itm_desc" type="text" id="std_itm_desc" value="<%=std_itm_desc%>" size="30" maxlength="100" />
                            <input type="hidden" name="st_no" id="st_no" value="<%=st_no%>" />
                            <input type="hidden" name="std_id" id="std_id" value="<%=std_id%>" />
                            </td>
                          <td align="right" bgcolor="#666666"><label for="std_unitcost"></label>
                            <input name="std_unitcost" type="text" id="std_unitcost" value="<%=std_unitcost%>" style="text-align: right;" size="12" maxlength="15" onkeydown="calctotal(document.formdodetail.std_unitcost.value, document.formdodetail.std_qty.value, document.formdodetail.std_subtotal);" onkeyup="calctotal(document.formdodetail.std_unitcost.value, document.formdodetail.std_qty.value, document.formdodetail.std_subtotal);" /></td>
                          <td align="right" bgcolor="#666666"><input name="std_qty" type="text" id="std_qty" style="text-align: right;" value="<%=std_qty%>" size="4" maxlength="10" onkeydown="calctotal(document.formdodetail.std_unitcost.value, document.formdodetail.std_qty.value, document.formdodetail.std_subtotal);" onkeyup="calctotal(document.formdodetail.std_unitcost.value, document.formdodetail.std_qty.value, document.formdodetail.std_subtotal);" /></td>
                          <td align="right" bgcolor="#666666"><input name="std_subtotal" type="text" id="std_subtotal" value="<%=std_subtotal%>" style="text-align: right;" size="12" maxlength="15" onkeydown="calctotal(document.formdodetail.std_unitcost.value, document.formdodetail.std_qty.value, document.formdodetail.std_subtotal);" onkeyup="calctotal(document.formdodetail.std_unitcost.value, document.formdodetail.std_qty.value, document.formdodetail.std_subtotal);" /></td>
                          <td align="right" bgcolor="#666666"><input type="button" name="button2" id="button2" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=request("dod_id")%>','action.asp?type=<%=stype%>','<%=dod_subtotal%>');" /></td>
                        </tr>
                        </form>
                         <%end if%> 
                         
 <%				i = 1
				sql1 = "SELECT std_id, std_st_no, std_itm_code, std_itm_desc, std_unitcost, std_qty, std_subtotal, std_referid " & _
	                  "FROM tblstockin_detail where std_st_no = '" & st_no & "' order by std_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
                         
                        <tr valign="top">
                          <td align="center"><%=i%>.</td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("std_itm_code")%></td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("std_itm_desc")%></td>
                          <td align="right" bgcolor="#FFFFFF" class='tktTotals'><%=chknumber2(rs1("std_unitcost"))%></td>
                          <td align="right" bgcolor="#FFFFFF" class='tktTotals'><%=rs1("std_qty")%></td>
                          <td align="right" bgcolor="#FFFFFF" class='tktTotals'><%=chknumber2(rs1("std_subtotal"))%></td>
                          <td align="right" nowrap="nowrap" bgcolor="#FFFFFF" class='tktTotals'><%if st_status="Open" then %>
                            <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_stockin_new.asp?std_id=<%=rs1("std_id")%>&amp;st_no=<%=rs1("std_st_no")%>#spareparts'" />
                            <input type="button" name="button9" id="button23" value="Del" onclick="javascript:confirmAction('<%=rs1("std_itm_code")%>','action.asp?type=delStockInDetail&std_id=<%=rs1("std_id")%>&st_no=<%=rs1("std_st_no")%>')" />
                           
                            <%end if%></td>
                        </tr>
 <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%> 
                        
                        <tr valign="top">
                          <td colspan="4" align="right" bgcolor="#FFFFFF"><strong> Total :</strong></td>
                          <td align="right" bgcolor="#FFFFFF" class='tktTotals'><strong><%=st_totalqty%></strong></td>
                          <td align="right" bgcolor="#FFFFFF" class='tktTotals'><strong><%=chknumber2(st_totalaAmt)%></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#FFFFFF" class='tktTotals'>&nbsp;</td>
                        </tr>
                        <tr>
                          <td colspan="7" valign="top"><table width="100%">
                            <tr>
                              <td width="50%" align="left" valign="middle" bgcolor="#FFFFFF"><%if st_status<>"Approved" and st_status<>"Cancel" then %>
                                <input type="button" name="DoneInvoice2" id="DoneInvoice2" value="Cancel Stock-In" onclick="javascript:confirmAction('<%=st_no%>','action.asp?type=CancelStockIn&amp;st_no=<%=st_no%>')" />
                                <%end if%>
                                <br />
                                <br /></td>
                              <td align="right" valign="middle" bgcolor="#FFFFFF"><%if st_status="Open" then %>
                              <input type="button" name="SubmitJob2" id="SubmitJob2" value="Submit Stock-In " onclick="javascript:confirmAction('<%=st_no%>','action.asp?type=SubmitStockIn&amp;st_no=<%=st_no%>')" />
                               <%end if%>
                                <br />
                                <%if st_status="Submitted" then %>
                                    <%if Request.Cookies("GAPS")("approve_stk") = "Y" then %>
                                        <input type="button" name="DoneInvoice" id="DoneInvoice" value="Approve Stock-In" onclick="javascript:confirmAction('<%=st_no%>','action.asp?type=ApproveStockIn&amp;st_no=<%=st_no%>')" />
                                     <%else%>
                                        <input type="button" name="DoneInvoice" id="DoneInvoice3" value="Approve - Not Authorized" disabled" />    
                                     <%end if%>
                                  <%end if%></td>
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
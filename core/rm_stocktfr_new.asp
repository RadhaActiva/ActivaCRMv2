<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            width: 414px;
        }
    </style>
</head>
<%

set rs = server.CreateObject("adodb.recordset")

if request("sf_no") <> "" then	  
sql = "SELECT sf_id, sf_no, sf_date, sf_referenceno, sf_status, sf_fromwarehouse, sf_towarehouse, sf_remark, sf_createddate, sf_createdby, sf_submitteddate, " & _
		"sf_submittedby, sf_approveddate, sf_approveddate, sf_approvedby, sf_cancelleddate, sf_cancelledby, sf_totalqty, sf_totalaAmt, sf_emailsent, sf_emailsentdate,sf_couriercompany " & _
		"FROM tblstocktransfer WHERE sf_no = '" & request("sf_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sf_id = rs("sf_id") 
			sf_no = rs("sf_no") 
			sf_date = rs("sf_date")
			sf_referenceno = rs("sf_referenceno") 
			sf_status = rs("sf_status")
			sf_fromwarehouse = rs("sf_fromwarehouse") 
			sf_towarehouse = rs("sf_towarehouse")
			sf_remark = rs("sf_remark")
			sf_createddate = rs("sf_createddate") 
			sf_createdby = rs("sf_createdby") 
			sf_submitteddate = rs("sf_submitteddate") 
			sf_submittedby = rs("sf_submittedby") 
			sf_approveddate = rs("sf_approveddate") 
			sf_approvedby = rs("sf_approvedby") 
			sf_cancelleddate = rs("sf_cancelleddate") 
			sf_cancelledby = rs("sf_cancelledby") 
			sf_totalqty = rs("sf_totalqty") 
			sf_totalaAmt = rs("sf_totalaAmt") 
			sf_emailsent = rs("sf_emailsent")  
			sf_emailsentdate = rs("sf_emailsentdate")
            sf_couriercompany=rs("sf_couriercompany")
		End If
		rs.Close
	  stype = "editStockTransfer"	
	  actionname = "Save" 
 else    
	  stype = "addStockTransfer"
	  actionname = "Save" 		
	  sf_date = date()	 
	  sf_status = "Open"   	
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Create/Edit </font>Stock Transfer</div></td>
                        <td align="right" class="titleblue1"><a href="rm_stocktfr_new_print.asp?sf_no=<%=sf_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/></a></td>
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
                        <td colspan="2" nowrap="nowrap" bgcolor="#DADADA"><strong>Stock-Transfer Information </strong></td>
                      </tr>
                    </thead>
                    
                    <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                      <tr class="head_row">
                        <td colspan="2" valign="top"><table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#EBEBEB">
                          <tr>
                            <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Stock-Transfer No</strong><br />
                             <strong><font size="1">(System Generate) </font></strong></font></td>
                            <td nowrap="nowrap"><strong><%=sf_no%><input name="sf_no" type="hidden" id="sf_no" value="<%=sf_no%>" />
                            </strong></td>
                            <td height="22" nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Stock-Transfer date</strong></font></strong></td>
                            <td><strong>
                              <input name="sf_date" type="text" id="sf_date" value="<%=chkdate(sf_date)%>" size="15" maxlength="20" />
                              <font color="#000000"><strong><a href="javascript:void(null)" onclick="window.dateField = document.formorder.do_purchase_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong></td>
                            <td bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>From Store</strong></font></strong></td>
                            <td><input name="sf_fromwarehouse" type="text" id="stkIn_warehouse" value="<%=sf_fromwarehouse%>" size="25" maxlength="50" />
                              [<a href="javascript:popup('rm_stockin_new_warehouse.asp?searchitem=wh_code&amp;searchvalue=<%=cusf_code%>&formname=formorder&fieldname=sf_fromwarehouse','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>]</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Reference No</strong></font></strong></td>
                            <td nowrap="nowrap"><input name="sf_referenceno" type="text" id="sf_referenceno" value="<%=sf_referenceno%>" size="20" maxlength="100" /></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong> Status</strong></font></strong></td>
                            <td align="left" nowrap="nowrap"><%=sf_status%></td>
                            <td align="left" nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>To Store</strong></font></strong></td>
                            <td align="left" nowrap="nowrap"><input name="sf_towarehouse" type="text" id="stkIn_warehouse2" value="<%=sf_towarehouse%>" size="25" maxlength="50" />
[<a href="javascript:popup('rm_stockin_new_warehouse.asp?searchitem=wh_code&amp;searchvalue=<%=cusf_code%>&amp;formname=formorder&amp;fieldname=sf_towarehouse','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>]</td>
                          </tr>
                            <tr>
                                <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Courrier Company</strong></font></td>
                                <td nowrap="nowrap">
                                <%if request.Cookies("GAPS")("slevel") = "technician" then %>
                                    <%=sp_couriercompany%>
                                    <input name="sf_couriercompany" type="hidden" id="sf_couriercompany" value="<%=sf_couriercompany%>" maxlength="30" />
                                <%else%>
                                  <select name="sf_couriercompany" id="sf_couriercompany">
                                  <option value="<%=sf_couriercompany%>"><%=sf_couriercompany%></option>
						          <option value="Self-Pickup">Self-Pickup</option>
                                  <!--<option value="Poslaju">Poslaju</option>-->
						          <option value="Nationwide">Nationwide</option>
                                  <option value="GDEX">GDEX</option>
                                  <option value="City-Link">City-Link</option>
                                  <option value="Skynet">Skynet</option>
                                </select>
                            <%end if%>
                        </td>
                            </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Created by</strong></font></strong></td>
                            <td nowrap="nowrap"><%=sf_createdby%> @ <%=chkdatetime(sf_createddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Remark</strong></font></strong></td>
                            <td colspan="3"><input name="sf_remark" type="text" id="supp_code" value="<%=sf_remark%>" size="50" maxlength="150" /></td>
                            </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Submitted By</strong></font></strong></td>
                            <td nowrap="nowrap"><%=sf_submittedby%> @ <%=chkdatetime(sf_submitteddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Cancelled by</strong></font></strong></td>
                            <td><%=sf_cancelledby%> @ <%=chkdatetime(sf_cancelleddate)%></td>
                            <td colspan="2" align="right">&nbsp;</td>
                          </tr>
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Approved By</strong></font></strong></td>
                            <td nowrap="nowrap"><%=sf_approvedby%> @ <%=chkdatetime(sf_approveddate)%></td>
                            <td nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td colspan="2" align="right"><%if sf_status<>"Approved" and sf_status<>"Cancel" then %><input type="submit" name="button" id="button" value="<%=actionname%>" />
                              <%end if%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr class="head_row">
                        <td colspan="2" valign="top">&nbsp;</td>
                      </tr>
                    </form>
                    

 
                    <%if sf_no <> "" then %>
                    
<%

if request("sfd_id") <> "" then
		sql = "SELECT sfd_id, sfd_st_no, sfd_itm_code, sfd_itm_desc, sfd_unitcost, sfd_qty, sfd_subtotal, sfd_referid,sfd_ex_qty " & _
	          "FROM tblstocktran_detail where sfd_id = '" & request("sfd_id") & "'"	
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   sfd_id = rs("sfd_id")
		   sfd_sf_no = rs("sfd_st_no")
		   sfd_itm_code = rs("sfd_itm_code")
		   sfd_itm_desc = rs("sfd_itm_desc")
		   sfd_unitcost = rs("sfd_unitcost")
		   sfd_qty = rs("sfd_qty")
		   sfd_subtotal = rs("sfd_subtotal")
		   sfd_referid = rs("sfd_referid")
           sfd_ex_qty = rs("sfd_ex_qty")
        end if
		rs.close
		sbutton = "Update"
		stype="editStockTransferDetail"	
else
		sbutton = "Add"
		stype="addStockTransferDetail"
		sfd_qty = "1"
        sfd_ex_qty ="0"
		sfd_unitcost = "0.00"	
		sfd_subtotal = "0.00"
end if

%>
                    <tr class="head_row">
                      <td colspan="2"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E8E8E8">
                      <tr class="head_row">
                          <td width="3%" height="24" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>No</strong></font></strong></td>
                          <td width="27%" align="left" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Item Code</strong></font></strong></td>
                          <td width="35%" align="left" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Item Name / Description</strong></font></strong></td>
                          <td width="15%" align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Stock Transfer Qty</strong></font></strong></td>
                          <td width="15%" align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Existing Qty</strong></font></strong></td>
                          <td width="20%" align="right" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Action</strong></font></strong></td>
                        </tr>
                        
                       <%if sf_status="Open" then %>
                       <form name="formdodetail" id="formdodetail" method="post" action="rm_jobsheet.asp#spareparts" >
                        <tr class="head_row">
                          <td height="24" colspan="2" bgcolor="#666666"><input name="sfd_itm_code" type="text" id="sfd_itm_code" value="<%=sfd_itm_code%>" size="25" maxlength="50" />
                            <a href="javascript:popup('rm_stocktfr_new_item.asp?searchitem=md_code&searchvalue=&sf_fromwarehouse='+document.formorder.sf_fromwarehouse.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">[Select]</a></td>
                          <td align="left" bgcolor="#666666"><input name="sfd_itm_desc" type="text" id="sfd_itm_desc" value="<%=sfd_itm_desc%>" size="30" maxlength="100" class="auto-style1" />
                            <input type="hidden" name="sf_no" id="sf_no" value="<%=sf_no%>" />
                            <input type="hidden" name="sfd_id" id="sfd_id" value="<%=sfd_id%>" />
                            <input type="hidden" name="sf_towarehouse" id="sf_towarehouse" value="<%=sf_towarehouse%>" />
                            </td>
                          <td align="center" bgcolor="#666666"><input name="sfd_qty" type="text" id="sfd_qty" style="text-align: right;" value="<%=sfd_qty%>" size="4" maxlength="10" /></td>
                          <td align="center" bgcolor="#666666"><%=sfd_ex_qty%></td>
                          <td align="center" bgcolor="#666666"><input type="button" name="button2" id="button2" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=request("dod_id")%>','action.asp?type=<%=stype%>','<%=dod_subtotal%>');" /></td>
                        </tr>
                        </form>
                         <%end if%> 
                         
 <%				i = 1
				sql1 = "SELECT sfd_id, sfd_st_no, sfd_itm_code, sfd_itm_desc, sfd_unitcost, sfd_qty, sfd_subtotal, sfd_referid, sfd_ex_qty " & _
	                  "FROM tblstocktran_detail where sfd_st_no = '" & sf_no & "' order by sfd_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
                         
                        <tr valign="top">
                          <td align="center"><%=i%>.</td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("sfd_itm_code")%></td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("sfd_itm_desc")%></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><%=rs1("sfd_qty")%></td>
                          <td align="center" bgcolor="#FFFFFF" ><%=rs1("sfd_ex_qty")%></td>
                          <td align="center" nowrap="nowrap" bgcolor="#FFFFFF" class='tktTotals'><%if sf_status="Open" then %>
                            <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_stocktfr_new.asp?sfd_id=<%=rs1("sfd_id")%>&amp;sf_no=<%=rs1("sfd_st_no")%>#spareparts'" />
                            <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmAction('<%=rs1("sfd_itm_code")%>','action.asp?type=delStockTransferDetail&sfd_id=<%=rs1("sfd_id")%>&sf_no=<%=rs1("sfd_st_no")%>')" />
                            <%end if%></td>
                        </tr>
 <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%> 
                        
                        <tr valign="top">
                          <td colspan="3" align="right" bgcolor="#FFFFFF"><strong> Total Qty:</strong></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><strong><%=sf_totalqty%></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#FFFFFF" class='tktTotals'>&nbsp;</td>
                        </tr>
                        <tr>
                          <td colspan="5" valign="top"><table width="100%">
                            <tr>
                              <td width="50%" valign="top" bgcolor="#FFFFFF"><%if sf_status<>"Approved" and sf_status<>"Cancel" then %>
                                <input type="button" name="DoneInvoice2" id="DoneInvoice2" value="Cancel Stock-Transfer" onclick="javascript:confirmAction('<%=sf_no%>','action.asp?type=CancelStockTransfer&amp;sf_no=<%=sf_no%>')" />
                                <%end if%>
                                <br />
                                <br /></td>
                              <td align="right" valign="top" bgcolor="#FFFFFF"><%if sf_status="Open" then %>
                                <input type="button" name="SubmitJob2" id="SubmitJob2" value="Submit Stock-Transfer " onclick="javascript:confirmAction('<%=sf_no%>','action.asp?type=SubmitStockTransfer&amp;sf_no=<%=sf_no%>')" />
                                <%end if%>
                                <br />
                                <%if sf_status="Submitted" then %>
                                   <%if Request.Cookies("GAPS")("approve_stk") = "Y" then %>
                                     <input type="button" name="DoneInvoice" id="DoneInvoice" value="Approve Stock-Transfer" onclick="javascript:confirmAction('<%=sf_no%>','action.asp?type=ApproveStockTransfer&amp;sf_no=<%=sf_no%>')" />
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
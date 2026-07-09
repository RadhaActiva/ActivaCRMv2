<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

if request("sp_no") <> "" then	  
sql = "SELECT sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
		"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
		"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
		"sp_rejecteddate, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
		"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate, sp_posteddate, sp_postedby, sp_canceldate, sp_cancelby " & _
		"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sp_id = rs("sp_id") 
			sp_no = rs("sp_no") 
			sp_tech_code = rs("sp_tech_code")
			sp_tech_name = rs("sp_tech_name")
			sp_tech_address = rs("sp_tech_address")
			sp_tech_postcode = rs("sp_tech_postcode") 
			sp_tech_state = rs("sp_tech_state") 
			sp_tech_city = rs("sp_tech_city") 
			sp_tech_email = rs("sp_tech_email") 
			sp_tech_tel1 = rs("sp_tech_tel1") 
			sp_tech_tel2 = rs("sp_tech_tel2") 
			sp_tech_carplateno = rs("sp_tech_carplateno") 
			sp_createddate = rs("sp_createddate")  
			sp_createdby = rs("sp_createdby")
			sp_date = rs("sp_date") 
			sp_status = rs("sp_status") 
			sp_submitteddate = rs("sp_submitteddate") 
			sp_submittedby = rs("sp_submittedby") 
			sp_approveddate = rs("sp_approveddate") 
			sp_approvedby = rs("sp_approvedby") 
			sp_deliverydate = rs("sp_deliverydate") 
			sp_deliveryby = rs("sp_deliveryby") 
			sp_trackingno = rs("sp_trackingno") 
			sp_couriercompany = rs("sp_couriercompany")
			sp_confirmedreceiveddate = rs("sp_confirmedreceiveddate") 
			sp_confirmedreceivedby = rs("sp_confirmedreceivedby") 
			sp_rejecteddate = rs("sp_rejecteddate") 
			sp_rejectedremark = rs("sp_rejectedremark") 
			sp_remark = rs("sp_remark") 
			sp_totalqty = rs("sp_totalqty") 
			sp_labourAmt = rs("sp_labourAmt") 
			sp_transportAmt = rs("sp_transportAmt") 
			sp_gstAmt = rs("sp_gstAmt") 
			sp_gstRate = rs("sp_gstRate") 
			sp_gstCode = rs("sp_gstCode") 
			sp_totalAmt = rs("sp_totalAmt") 
			sp_emailsent = rs("sp_emailsent") 
			sp_emailsentdate = rs("sp_emailsentdate") 
			sp_logby = rs("sp_logby") 
			sp_logdate = rs("sp_logdate") 
			sp_postedby = rs("sp_postedby") 
			sp_posteddate = rs("sp_posteddate") 
			sp_cancelby = rs("sp_cancelby") 
			sp_canceldate = rs("sp_canceldate") 
		End If
		rs.Close
	  stype = "editSparepartsRequest"	
	  actionname = "Save" 
 else    
	  stype = "AddSparepartsRequest"
	  actionname = "Save" 		
	  sp_date = date()	 
	  sp_status = "Open"   	
	  

      sql = "SELECT tech_id, tech_code, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
            "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id " & _
	        "FROM tbltechnician WHERE tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sp_tech_code = rs("tech_code")
			sp_tech_name = rs("tech_name")
			sp_tech_icno = rs("tech_icno")
			sp_tech_address = rs("tech_address")
			sp_tech_postcode = rs("tech_postcode") 
			sp_tech_state = rs("tech_state") 
			sp_tech_state_id = rs("tech_state_id") 
			sp_tech_city = rs("tech_city") 
			sp_tech_city_id = rs("tech_city_id") 
			sp_tech_email = rs("tech_email") 
			sp_tech_tel1 = rs("tech_tel1") 
			sp_tech_tel2 = rs("tech_tel2") 
			sp_tech_carmodel = rs("tech_carmodel") 
			sp_tech_carplateno = rs("tech_carplateno") 
			sp_tech_carcolour = rs("tech_carcolour") 
		End If
		rs.Close
			  
end if
 
%>

<script language="javascript">

function confirmForm(id,orderlinks,otype) 
{
	
  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formorderparts.action = orderlinks;
	document.formorderparts.submit();
   }
}


function calctotal(unitprice,qty,subtotal) {

    var temp = 0;
	temp = unitprice * qty;
	subtotal.value = temp.toFixed(2);

}
// -->
</script>

        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="77%" class="titleblue1"><div align="left"><font color="#CC0000">Create </font>Spare Part Request</div></td>
                        <td width="23%" align="right" class="titleblue1"><a href="rm_consignment_print.asp?sp_no=<%=sp_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/></a></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
                
                <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                <tr>
                  <td width="49%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                        <td colspan="4" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Technician  
                          Information </font></strong></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Techician Code*</strong></font></td>
                        <td colspan="3" align="left"><label for="textfield4"></label>
                          <strong>
                         <%if request.Cookies("GAPS")("slevel") = "technician" then %>
                         
						    <%=sp_tech_code%>
                            <input name="sp_tech_code" type="hidden" id="sp_tech_code" value="<%=sp_tech_code%>" />
                         <%else%>
                            
						 <select name="sp_tech_code" id="sp_tech_code" style="width: 200px">
                            <option value=""></option>
							<%			
                            sql = "SELECT tech_id, tech_code, tech_name, tech_wh_code FROM tbltechnician where tech_status = 'Y' and tech_wh_code<>'0' "	
                            set rs = server.CreateObject("adodb.recordset")
                            rs.Open sql,strconnect,3,3,&H0001
                            while Not rs.EOF
								if (sp_tech_code) = (rs("tech_code")) then
								response.write "<option value='" & rs("tech_code") & "' selected>" & rs("tech_code") & " - " & rs("tech_wh_code") & " - " & rs("tech_name")  & "</option>"
								else
								response.write "<option value='" & rs("tech_code") & "'>" & rs("tech_code") & " - " & rs("tech_wh_code") & " - " & rs("tech_name")  & "</option>"
								end if 					  
								rs.movenext
								wend
                            rs.close					
                            %>
               			 <%end if%>
                          </select>
                          
                          </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Techician Name*</strong></font></td>
                        <td colspan="3" align="left"><%=sp_tech_name%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                        <td colspan="3" align="left"><%=sp_tech_address%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td colspan="3" align="left"><%=sp_tech_postcode%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font></td>
                        <td width="33%" align="left"><%=sp_tech_state%></td>
                        <td width="13%" align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">City*</font></strong></td>
                        <td width="32%" align="left"><%=sp_tech_city%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td colspan="3" align="left"><label for="select8"><%=sp_tech_email%></label></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td colspan="3" valign="top"><%=sp_tech_tel1%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td colspan="3" valign="top"><label for="textfield3"><%=sp_tech_tel2%></label></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><strong><font color="#FFFFFF">Car Plate No</font></strong></td>
                        <td colspan="3" valign="top"><%=sp_tech_carplateno%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Created by</strong></font></td>
                        <td colspan="3" valign="top"><%=sp_createdby%> @ <%=chkdatetime(sp_createddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> Spare Part Information</font></strong></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Request  
                          No.<br />
                          <font size="1">(System Generate) </font></strong></font></td>
                        <td align="left" nowrap="nowrap"><strong><%=sp_no%></strong></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Requested Date</strong></font></td>
                        <td align="left" nowrap="nowrap"><%=chkdate(sp_date)%></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Status</strong></font></td>
                        <td align="left" nowrap="nowrap"><strong><%=sp_status%></strong></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Submitted Date </strong></font></td>
                        <td nowrap="nowrap"><%=sp_submittedby%> @ <%=chkdate(sp_submitteddate)%></td>
                      </tr>
                      <tr align="left" >
                        <td nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                        <td nowrap="nowrap"><label for="textfield5"></label></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Approved  Date</strong></font></td>
                        <td nowrap="nowrap"><%=sp_approvedby%> @ <%=chkdate(sp_approveddate)%></td>
                      </tr>
                      <tr>
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Courrier Track No.</strong></font></td>
                        <td align="left" nowrap="nowrap">
                        <%if request.Cookies("GAPS")("slevel") = "technician" then %>
                        <%=sp_trackingno%>
                        <input name="sp_trackingno" type="hidden" id="sp_trackingno" value="<%=sp_trackingno%>" maxlength="30" />
						<%else%>
                        <input name="sp_trackingno" type="text" id="sp_trackingno" value="<%=sp_trackingno%>" maxlength="30" />
                        <%end if%>
                        
                        </td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Delivery  Date</strong></font></td>
                        <td align="left" nowrap="nowrap"><label for="select2"><font color="#000000"><strong>
                          <%if request.Cookies("GAPS")("slevel") = "technician" then %>
                          <%=sp_deliveryby%> @ <%=chkdate(sp_deliverydate)%>
                          <%else%>
                          <input name="sp_deliverydate" type="text" id="sp_deliverydate" value="<%=chkdate(sp_deliverydate)%>" size="12" maxlength="20" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.formorder.sp_deliverydate;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font>
                          <%end if%>
             </label>
             </td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Courrier Company</strong></font></td>
                        <td nowrap="nowrap">
                        <%if request.Cookies("GAPS")("slevel") = "technician" then %>
                        <%=sp_couriercompany%>
                        <input name="sp_couriercompany" type="hidden" id="sp_couriercompany" value="<%=sp_couriercompany%>" maxlength="30" />
                        <%else%>
                          <select name="sp_couriercompany" id="sp_couriercompany">
                          <option value="<%=sp_couriercompany%>"><%=sp_couriercompany%></option>
						  <option value="Self-Pickup">Self-Pickup</option>
                          <option value="Poslaju">Poslaju</option>
						  <option value="Nationwide">Nationwide</option>
                        <!--  <option value="GDEX">GDEX</option>
                          <option value="City-Link">City-Link</option>
                          <option value="Skynet">Skynet</option>-->
                        </select>
                        <%end if%>
                        </td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Confirmed Received Date</strong></font></td>
                        <td nowrap="nowrap"><font color="#000000"><strong>
                        <%if request.Cookies("GAPS")("slevel") = "technician" then %>
                        <%=sp_confirmedreceivedby%> @ <%=chkdate(sp_confirmedreceiveddate)%>
                        <%else%>  
                        <input name="sp_confirmedreceiveddate" type="text" id="sp_confirmedreceiveddate" value="<%=chkdate(sp_confirmedreceiveddate)%>" size="12" maxlength="20" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.formorder.sp_confirmedreceiveddate;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font>
                         <%end if%>
             </td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                        <td nowrap="nowrap">&nbsp;</td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Posted Date </strong></font></td>
                        <td nowrap="nowrap"><%=sp_postedby%> @ <%=chkdate(sp_posteddate)%></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155">&nbsp;</td>
                        <td nowrap="nowrap"><label for="sp_couriercompany"></label></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Rejected  Date</strong></font></td>
                        <td nowrap="nowrap"><%=sp_rejectedby%> @ <%=chkdate(sp_rejecteddate)%></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                        <td colspan="3" nowrap="nowrap"><strong>
                          <textarea name="sp_remark" cols="50" rows="3" wrap="virtual" id="sp_remark"><%=sp_remark%></textarea>
                        </strong></td>
                      </tr>
                      <tr >
                        <td colspan="4" align="right" nowrap="nowrap"><input name="sp_no" type="hidden" id="sp_no" value="<%=sp_no%>" />
                        <%if request.Cookies("GAPS")("slevel") = "technician" then %>
                            <%if sp_status="Open" then %>
                             <input type="submit" name="button6" id="button6" value="Save" />
                            <%end if%>
                        <%else%>
                             <input type="submit" name="button6" id="button6" value="Save" />
                        <%end if%>     
                          </td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                </form>
                
<%if sp_no <> "" then  %>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
             
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#475387">
                          <td colspan="2"><font color="#FFFFFF"><strong>Spare Part 
                            Code</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong> Description</strong></font></td>
                          <td width="10%" align="center"><font color="#FFFFFF"><strong>My Stock Qty</strong></font></td>
                          <!--Added By sanjay on 23/Feb/2012-->
                          <td align="right"><font color="#FFFFFF"><strong>Unit Price (RM)</strong></font></td>
                          <td width="10%" align="right"><font color="#FFFFFF"><strong>Qty</strong></font></td>
                          <td align="right"><font color="#FFFFFF"><strong>Total 
                            Amt (RCP)</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Action</strong></font></td>
                        </tr>
 
 
 <%

if request("spd_id") <> "" then
		sql = "SELECT spd_id, spd_sp_no, spd_tech_code, spd_partcode, spd_currentstock, spd_desc, spd_unitcost, spd_qty, spd_subtotal " & _
	           "FROM tblsparepartrequest_detail where spd_id = '" & request("spd_id") & "' "	 
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   spd_id = rs("spd_id")
		   spd_sp_no = rs("spd_sp_no")
		   spd_tech_code = rs("spd_tech_code")
		   spd_partcode = rs("spd_partcode")
		   spd_currentstock = rs("spd_currentstock")
		   spd_desc = rs("spd_desc")
		   spd_unitcost = rs("spd_unitcost")
		   spd_qty = rs("spd_qty")
		   spd_subtotal = rs("spd_subtotal")
        end if
		rs.close
		sbutton = "Update"
		stype="editSparepartsRequestDetail"	
else
		sbutton = "Add"
		stype="addSparepartsRequestDetail"
		spd_qty = "1"	
		spd_unitcost = "0.00"	
		spd_subtotal = "0.00"	
end if

%>
                       <%if sp_status="Open" then %> 
                        <form name="formorderparts" id="formorderparts" method="post" action="action.asp?type=<%=stype%>" >
                        <tr>
                          <td colspan="2" bgcolor="#666666"><input name="spd_partcode" type="text" id="spd_partcode" value="<%=spd_partcode%>" maxlength="50" />
[<a href="javascript:popup('rm_consignment_new_spareparts.asp?searchitem=md_code&amp;searchvalue=<%=cust_code%>&amp;formname=formorderparts&amp;fieldname=spd_partcode&sp_type=All','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td>
                          <td align="center" bgcolor="#666666"><font color="#FFFFFF"> </font>
                            <input name="spd_desc" type="text" id="spd_desc" value="<%=spd_desc%>" maxlength="100" /></td>
                          <td align="center" bgcolor="#666666"><input name="spd_id" type="hidden" id="spd_id" value="<%=spd_id%>" />
                            <input name="sp_no" type="hidden" id="sp_no" value="<%=sp_no%>" />
                            <a name="spareparts" id="spareparts"></a></td>
                          <td align="right" bgcolor="#666666"><font color="#FFFFFF">
                            <input name="spd_unitcost" type="text" id="spd_unitcost" style="text-align:right; background-color: #cccccc;" onfocus="this.blur();" onkeydown="calctotal(document.formorderparts.spd_unitcost.value, document.formorderparts.spd_qty.value, document.formorderparts.spd_subtotal);" onkeyup="calctotal(document.formorderparts.spd_unitcost.value, document.formorderparts.spd_qty.value, document.formorderparts.spd_subtotal);" value="<%=spd_unitcost%>" size="5" maxlength="10" />
                          </font></td>
                          <td align="right" bgcolor="#666666"><input name="spd_qty" type="text" id="spd_qty" style="text-align:right" onkeydown="calctotal(document.formorderparts.spd_unitcost.value, document.formorderparts.spd_qty.value, document.formorderparts.spd_subtotal);" onkeyup="calctotal(document.formorderparts.spd_unitcost.value, document.formorderparts.spd_qty.value, document.formorderparts.spd_subtotal);" value="<%=spd_qty%>" size="5" maxlength="5" /></td>
                          <td align="right" bgcolor="#666666"><input name="spd_subtotal" type="text" id="spd_subtotal" style="text-align:right; background-color: #cccccc;" onfocus="this.blur();" value="<%=spd_subtotal%>" size="10" maxlength="10" /></td>
                          <td align="center" bgcolor="#666666"><input type="submit" name="additems33" id="additems33" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=request("jobp_id")%>','action.asp?type=<%=stype%>','<%=jobp_subtotal%>');" /></td>
                        </tr>
                        </form>
                        <%end if%>
                        
<%				i = 1
				sql1 = "SELECT spd_id, spd_sp_no, spd_tech_code, spd_partcode, spd_currentstock, spd_desc, spd_unitcost, spd_qty, spd_subtotal " & _
	                   "FROM tblsparepartrequest_detail where spd_sp_no = '" & sp_no & "' order by spd_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%> 
                        <tr>
                          <td align="center"><%=i%>.</td>
                          <td align="left"><%=rs1("spd_partcode")%></td>
                          <td align="left"><%=rs1("spd_desc")%></td>
                          <td align="center"><%=rs1("spd_currentstock")%></td>
                          <td align="right"><%=rs1("spd_unitcost")%></td>
                          <td align="right"><%=rs1("spd_qty")%></td>
                          <td align="right"><%=chknumber2(rs1("spd_subtotal"))%></td>
                          <td align="center" nowrap="nowrap">
						  <%if sp_status="Open" then %>
                            <input type="button" name="aditem211" id="aditem211" value="Edit" onclick="document.location.href='rm_consignment_new.asp?spd_id=<%=rs1("spd_id")%>&amp;sp_no=<%=sp_no%>#spareparts'" />
                            <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmDel('<%=rs1("spd_partcode")%>','action.asp?type=DelSparepartsRequestDetail&spd_id=<%=rs1("spd_id")%>&sp_no=<%=sp_no%>')" />
                            <%end if%></td>
                        </tr>
 <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="3" align="right"><strong>Total</strong><div class="total1"> </div></td>
                          <td align="center">&nbsp;</td>
                          <td height="25" align="right">&nbsp;</td>
                          <td height="25" align="right"><%=sp_totalqty%></td>
                          <td align="right"> <strong><%=chknumber2(sp_totalAmt)%></strong></td>
                          <td>&nbsp;</td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <!--<td align="left" bgcolor="#FFFFFF" 
          scope="col"><p>
                        
                        Email
                            <input name="textfield9" type="text" id="textfield11" size="50" />
  <input type="submit" name="button4" id="button4" value="Resend Email Request " />
  <br />
                        <br />
                        <br />
                        <br />
                        <strong><br />
                        </strong></p></td>-->
                      <td align="right" bgcolor="#FFFFFF" 
          scope="col">
          <%if request.Cookies("GAPS")("slevel") = "technician" then %>
              <%if sp_status="Open" then %>
              <input type="button" name="button5" id="button5" value="Submit Request" onclick="javascript:confirmAction('<%=sp_no%>','action.asp?type=SubmitSparepartsRequest&sp_no=<%=sp_no%>')" />
              <%end if%>
		  
		  <%elseif request.Cookies("GAPS")("slevel") = "sc" or request.Cookies("GAPS")("slevel") = "cs" then %>
              <%if sp_status="Submitted" then %>
                  <input type="button" name="spbutton" id="spbutton" value="Approve Request" onclick="javascript:confirmAction('<%=sp_no%>','action.asp?type=ApproveSparepartsRequest&sp_no=<%=sp_no%>')" />
                  <input type="button" name="spbutton" id="spbutton" value="Reject Request" onclick="javascript:confirmAction('<%=sp_no%>','action.asp?type=RejectSparepartsRequest&sp_no=<%=sp_no%>')" />
              <%elseif sp_status="Approved" then %>    
                  <input type="button" name="spbutton" id="spbutton" value="Delivered Request" onclick="javascript:confirmAction('<%=sp_no%>','action.asp?type=DeliveredSparepartsRequest&sp_no=<%=sp_no%>')" />
               <%elseif sp_status="Delivered" then %>  
                  
                  <%if left(Ucase(sp_tech_code),2) <> "IP" then  %>  
                  <input type="button" name="spbutton" id="spbutton" value="Posted Request" onclick="javascript:confirmAction('<%=sp_no%>','action.asp?type=PostedSparepartsRequest&sp_no=<%=sp_no%>')" />    
                  <%end if%>
			   <%end if%>  
               
               <br><br>
               
               <%if request.Cookies("GAPS")("slevel") = "sc" or request.Cookies("GAPS")("slevel") = "cs" then %>
				   <%if sp_status="Open" then %>
                   <input type="button" name="button5" id="button5" value="Submit Request" onclick="javascript:confirmAction('<%=sp_no%>','action.asp?type=SubmitSparepartsRequest&sp_no=<%=sp_no%>')" />
                   <%elseif sp_status<>"Posted" then%>
                   <input type="button" name="spbutton" id="spbutton" value="Revert Request" onclick="javascript:confirmAction('<%=sp_no%>','action.asp?type=RevertSparepartsRequest&sp_no=<%=sp_no%>')" /> 
				   <%end if%>			   
			   <%end if%>
               
               
               <%if sp_status<>"Posted" then %>
               <input type="button" name="spbutton" id="spbutton" value="Cancel Request" onclick="javascript:confirmAction('<%=sp_no%>','action.asp?type=CancelSparepartsRequest&sp_no=<%=sp_no%>')" /> 
               <%end if%>
          <%end if%>
                   </td>
                    </tr>
                    <tr align="right">
                      <td colspan="2" bgcolor="#FFFFFF"></td>
                    </tr>
                    <tr>
                      <td></tbody></td>
                    </tr>
                  </table></td>
                </tr>
                
<%end if %>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
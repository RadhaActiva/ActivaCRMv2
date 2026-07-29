<!-- #include file="database/datastore.asp" -->

<%

jobmonth= request("jobmonth")
jobyear= request("jobyear")
tech_code = request("tech_code")

if request("copylink") = "Yes" and request("verified") <> "" then
    program_str=Request.ServerVariables("URL")
    query_str= Request.ServerVariables("QUERY_STRING")
    fullpath=program_str + "?" + query_str
    'fullpath=Replace(fullpath,"/baby/","")
    fullpath=Replace(fullpath,"&copylink=Yes","")    
    
    sql = "SELECT tech_code, clink, createdatetime FROM tbltech_claim_forms WHERE clink = '" & fullpath & "' and tech_code ='" & tech_code & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
    	rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if rs.eof then
		    rs.addnew
            rs("tech_code") = tech_code
		    rs("clink") = fullpath
            rs("createdatetime") = ChkDateTimeMySQL(now())
    		rs.Update 
	    	rs.Close 
        end if	 
end if

set rs = server.CreateObject("adodb.recordset")
set rs1 = server.CreateObject("adodb.recordset")

if request("rpc_id") <> "" then	  
sql2 = "SELECT rpc_id, rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_serviceQty1, rpc_serviceAmt1, rpc_serviceQty2, rpc_serviceAmt2, " & _
		"rpc_techfees, rpc_car_allow, rpc_phone_allow, rpc_toll, rpc_parking, rpc_petrol, rpc_hotel, rpc_service_allow, rpc_overwarranty_qty, rpc_overwarranty_fee, rpc_others,  " & _
		"rpc_deduction_ow_qty, rpc_deduction_ow, rpc_deduction_sparepart_qty, rpc_deduction_sparepart, rpc_deduction_desc1,rpc_deduction_desc2,rpc_deduction_total, rpc_total,rpc_others_desc,rpc_others_desc2, rpc_submitted_date, rpc_checkedby, rpc_checked_date, " & _
		"rpc_verifiedby, rpc_verified_date,rpc_others2,rpc_water_storage_qty " & _
		"FROM tblrpr_techcommission where rpc_id = " & request("rpc_id") & " "
		rs.Open sql2,strconnect,0,1,&H0001
		If Not rs.EOF Then
			rpc_id = rs("rpc_id") 
			rpc_month = rs("rpc_month") 
			rpc_year = rs("rpc_year") 
			rpc_tech_code = rs("rpc_tech_code") 
			rpc_tech_name = rs("rpc_tech_name")
			rpc_serviceQty1 = rs("rpc_serviceQty1")
			rpc_serviceAmt1 = rs("rpc_serviceAmt1") 
			rpc_serviceQty2 = rs("rpc_serviceQty2") 
			rpc_serviceAmt2 = rs("rpc_serviceAmt2") 
			rpc_techfees = rs("rpc_techfees") 
			rpc_car_allow = rs("rpc_car_allow") 
			rpc_phone_allow = rs("rpc_phone_allow") 
			rpc_toll = rs("rpc_toll") 
			rpc_parking = rs("rpc_parking") 
			rpc_petrol = rs("rpc_petrol") 
			rpc_hotel = rs("rpc_hotel") 
			rpc_service_allow = rs("rpc_service_allow") 
			rpc_overwarranty_qty = rs("rpc_overwarranty_qty")  
			rpc_overwarranty_fee = rs("rpc_overwarranty_fee") 
	
    		rpc_others = rs("rpc_others") 
            rpc_others_desc = rs("rpc_others_desc") 
            rpc_others2 = rs("rpc_others2") 
            rpc_others_desc2 = rs("rpc_others_desc2") 

			rpc_deduction_ow_qty = rs("rpc_deduction_ow_qty") 
			rpc_deduction_ow = rs("rpc_deduction_ow") 
			rpc_deduction_sparepart = rs("rpc_deduction_sparepart") 
			rpc_deduction_sparepart_qty = rs("rpc_deduction_sparepart_qty") 
			rpc_deduction_total = rs("rpc_deduction_total") 

            rpc_deduction_desc1 = rs("rpc_deduction_desc1") 
            rpc_deduction_desc2 = rs("rpc_deduction_desc2") 
			rpc_total = rs("rpc_total")
            rpc_submitted_date   = rs("rpc_submitted_date")
            rpc_checkedby =  rs("rpc_checkedby")
            rpc_checked_date = rs("rpc_checked_date")
            rpc_verifiedby = rs("rpc_verifiedby")
            rpc_verified_date = rs("rpc_verified_date")
            rpc_water_storage_qty = rs("rpc_water_storage_qty")
		End If
		rs.Close

            sql3 = "select sum(te_net_mileage) from tbltech_claim_exmileage where te_month =  " & jobmonth & " and te_year = " & jobyear & " and te_tech_code ='" & tech_code & "'"
        	total_km = selectid(sql3)
end if


if rpc_tech_code <> "" then	  
sql = "SELECT tech_id, tech_code, tech_type, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
      "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id, tech_wh_code, tech_salary,b.state_name " & _
	  "FROM tbltechnician a join tblstate b on b.state_id = a.tech_area_id WHERE a.tech_code = '" & rpc_tech_code & "' "
		rs1.Open sql,strconnect,0,1,&H0001
		If Not rs1.EOF Then
			tech_id = rs1("tech_id") 
			tech_code = rs1("tech_code") 
			tech_type = rs1("tech_type") 
			tech_name = rs1("tech_name") 
			tech_icno = rs1("tech_icno")
			tech_address = rs1("tech_address")
			tech_postcode = rs1("tech_postcode") 
			tech_state = rs1("tech_state") 
			tech_state_id = rs1("tech_state_id") 
			tech_city = rs1("tech_city") 
			tech_city_id = rs1("tech_city_id") 
			tech_email = rs1("tech_email") 
			tech_tel1 = rs1("tech_tel1") 
			tech_tel2 = rs1("tech_tel2") 
			tech_createdby = rs1("tech_createdby") 
			tech_cretateddate = rs1("tech_cretateddate") 
			tech_carmodel = rs1("tech_carmodel")  
			tech_carplateno = rs1("tech_carplateno") 
			tech_carcolour = rs1("tech_carcolour") 
			tech_password = rs1("tech_password") 
			tech_status = rs1("tech_status") 
			tech_area = rs1("state_name")
			tech_area_id = rs1("tech_area_id")
			tech_wh_code = rs1("tech_wh_code")

            set rs11 = server.CreateObject("adodb.recordset")
            sql11 = "SELECT post_office from tblpostcode WHERE postcode = '" & tech_postcode & "' "
            rs11.Open sql11,strconnect,0,1,&H0001   
            If Not rs11.EOF Then
                tech_area = Trim(rs11("post_office") & "")
                If Trim(rs1("tech_state") & "") <> "" Then
                    If tech_area <> "" Then
                        tech_area = tech_area & ", "
                    End If
                    tech_area = tech_area & Trim(rs1("tech_state") & "")
                End If
            End If
            rs11.close
            end if
End If
rs1.Close

if rpc_total = "" then
    response.write ("Error ! Please contact Service Coordinator")
    response.end
end if

%>
<script language="javascript">

function ApproveClaim(id,orderlinks,otype) 
{
    
   // if (document.forms["form1"].checkedOK.checked = true)
     //   alert("checked");

    document.form1.action = orderlinks;
    document.form1.submit()
    }
</script>

<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>
<form name="form1" id="form1" method="post" action="action.asp?type=<%=stype%>&tech_code=<%=tech_code%>&techtype=TPC">
 <table width="900" border="0" cellspacing="0" cellpadding="3">
  <tr>
    <th scope="row"><img src="images/Riegen1.png" width="100" height="100"><br><br>
      Riegen Marketing Sdn Bhd <small>202401008163 (1554013-U)</small><br>
      B-3-A-18 & B-3A-19, Block Bougainvilla, 0 Boulevard, Lebuhraya Sprint, PJU6A<br>
      47400 Petaling Jaya, Selangor Darul Ehsan <br />
      Tel:  03-77319139<br>
      Website : www.riegen.com.my</th>
  </tr>
  <tr>
    <th scope="row">&nbsp;</th>
  </tr>
  <tr>
    <th scope="row"><p><u><h2>CLAIM FORM</h2></u></p>
    <p>3rd PARTY CONTRACTOR (TPC)</p></th>
  </tr>
  <tr>
    <th scope="row">&nbsp;</th>
  </tr>
  <tr>
    <th scope="row"><table width="100%" border="1" cellspacing="0" cellpadding="3">
      <tr>
        <th width="24%" align="left" bgcolor="#E5E5E5" scope="row">Name of Company :</th>
        <th colspan="3" align="left" scope="row"><%=tech_name%></th>
        <th width="17%" align="center" bgcolor="#E5E5E5" scope="row">MM/YY of : <br></th>
        <th width="16%" align="center" scope="row"><%=convertmonth(rpc_month)%>/<%=rpc_year%></th>
      </tr>
      <tr>
        <th align="left" bgcolor="#E5E5E5" scope="row">Contrator Code : </th>
        <th width="15%" align="left" scope="row"><%=rpc_tech_code%></th>
        <th width="15%" align="center" bgcolor="#E5E5E5" scope="row">Area :</th>
        <th width="13%" align="center" scope="row"><%=tech_area%></th>
        <th align="center" bgcolor="#E5E5E5" scope="row">Date :</th>
        <th align="center" scope="row"><%=chkdate(date())%></th>
      </tr>
    </table></th>
  </tr>
  <tr>
    <th scope="row">&nbsp;</th>
  </tr>
  <tr>
    <th align="left" scope="row">A: Claims</th>
  </tr>
  <tr>
    <th scope="row"><table width="100%" border="1" cellspacing="0" cellpadding="3">
      <tr>
        <th align="left" width="200" bgcolor="#E5E5E5" scope="row">&nbsp;</th>
        <th align="left" bgcolor="#E5E5E5" scope="row">&nbsp;</th>
        <th width="50" align="center" bgcolor="#E5E5E5" scope="row">Qty</th>
        <th width="150" align="center" bgcolor="#E5E5E5" scope="row">Remark</th>
        <th width="80" align="center" bgcolor="#E5E5E5" scope="row"> Claim</th>
      </tr>
      <tr>
        <th align="left" scope="row">1. Service Incentives Ceiling Fan &amp; Water Heater<br>
          (Under Warranty Only)</th>
        <th align="left" scope="row">RM45 per cases<br></th>
        <th align="center" scope="row"><%=rpc_serviceQty1%></th>
        <th align="left" scope="row">&nbsp;</th>
        <th align="center" scope="row"><%=chknumber2(rpc_serviceAmt1)%></th>
      </tr>
      <tr>
        <th align="center" scope="row">2nd Unit</th>
        <td align="left" scope="row"><strong style="color: rgb(0, 0, 0); font-family: Tahoma; font-size: 12px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; letter-spacing: normal; orphans: 2; text-align: -webkit-left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">RM25 per Unit.</strong><br style="color: rgb(0, 0, 0); font-family: Tahoma; font-size: 12px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: -webkit-left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;" />
            <span style="color: rgb(0, 0, 0); font-family: Tahoma; font-size: 12px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: -webkit-left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;">(2nd units onwards under same address)</span></td>
        <th align="center" scope="row"><%=rpc_serviceQty2%></th>
        <th align="left" scope="row">&nbsp;</th>
        <th align="center" scope="row"><%=chknumber2(rpc_serviceAmt2)%></th>
      </tr>
      <tr>
        <td align="left" scope="row"><strong>2. Petrol Extra milege Claim </strong><br>
          (Please refer attached document) <br>
          <br></td>
        <td align="left" scope="row"><strong>(RM0.60/KM after 50km)</strong></td>
        <th align="center" scope="row"><%=total_km%> KM</th>
        <th align="left" scope="row">
             <% sql1 = "SELECT tc_fuel_receipt FROM tbltech_claim where tc_tech_code = '" & rpc_tech_code & "' and tc_year_process = '" & jobyear & "' and tc_month_process = '" & jobmonth & "'"	  
		       set rs1 = server.CreateObject("adodb.recordset")
		       rs1.Open sql1,strconnect,3,3,&H0001
               If Not rs1.EOF Then
                   if not isnull(rs1("tc_fuel_receipt")) then 
                        arr = split (rs1("tc_fuel_receipt"),",")
                        arrbound = ubound(arr)
                        for i = 0 to arrbound%>
                            <a href="shared/claims/<%=arr(i)%>" target="_blank"><%=arr(i)%></a>
                    <%    next
                   end if			
                end if
		       rs1.close
            %>
        </th>
        <th align="center" scope="row"><%=chknumber2(rpc_petrol)%></th>
      </tr>
  <tr>
        <!--<th align="left" scope="row">Storage Water Tank RM80/Case </th>
        <td align="left" scope="row"><%=rpc_others_desc%></td>
        <th align="center" scope="row"><%=rpc_water_storage_qty%></th>
        <th align="left" scope="row">&nbsp;</th>
        <th align="center" scope="row"><%=chknumber2(rpc_others)%></th>-->
      </tr>

       <tr>
        <th align="left" scope="row">&nbsp;Other b): </th>
        <td align="left" colspan="3" scope="row"><%=rpc_others_desc2%></td>
         <th align="center" scope="row"><%=chknumber2(rpc_others2)%></th>
      </tr>
    </table></th>
  </tr>
  <tr>
    <td scope="row">&nbsp;</td>
  </tr>
  <tr>
    <td scope="row"><strong>B: Deduction/Over Warranty</strong></td>
  </tr>
  <tr>
    <td scope="row"><table width="100%" border="1" cellspacing="0" cellpadding="3">
      <tr>
        <th align="left" bgcolor="#E5E5E5" scope="row">&nbsp;</th>
         <th width="50" align="center" bgcolor="#E5E5E5" scope="row">Qty</th>
        <th width="150" align="center" bgcolor="#E5E5E5" scope="row">Remark</th>
        <th width="80" align="center" bgcolor="#E5E5E5" scope="row"> RM </th>
      </tr>
      <tr>
        <td align="left" scope="row"><strong>1. Over Warranty Service Case &amp; Spare-Part Incentive:-</strong><br>
          (Please refer Attached Job Sheet Report)<br></td>
        <th align="center" scope="row"><%=rpc_overwarranty_qty%></th>
        <th align="left" scope="row">&nbsp;</th>
        <th align="center" scope="row"><%=chknumber2(rpc_overwarranty_fee)%></th>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td scope="row">&nbsp;</td>
  </tr>
    <tr>
        <td scope="row"><strong>C: Deduction</strong></td>
  </tr>
    <tr>
    <td scope="row"><table width="100%" border="1" cellspacing="0">
      <tr>
        <td align="left" scope="row"><%=rpc_deduction_desc1%><br></td>
        <th width="80" align="center" scope="row"><%=chknumber2(rpc_deduction_ow)%></th>
        </tr>
        <tr>
        <td align="left" scope="row"><%=rpc_deduction_desc2%><br></td>
        <th width="80" align="center" scope="row"><%=chknumber2(rpc_deduction_sparepart)%></th>
        </tr>
        </table>
  <tr>
    <td scope="row"><table width="100%" border="1" cellspacing="0" cellpadding="3">
      <tr>
        <td align="left" scope="row"><strong>TOTAL AMOUNT OF CLAIM </strong><br></td>
        <th width="80" align="center" scope="row"><%=chknumber2(rpc_total)%></th>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td scope="row">&nbsp;</td>
  </tr>
  <tr>
    <td scope="row"><strong>Remark :- </strong><br>
      1) This Monthly Service Incentives Claim will only be valid upon job completion with the attachment of complete Job Sheet (with warranty card/proof of puchase). <br>
2) All service jobs must completed within three(3) working days from the date of complaint. <br>
3) Riegen Marketing Sdn Bhd reserves the right to change or amend the incentive rate and regulations at any time without prior notice. </td>
  </tr>
  <tr>
    <td scope="row">&nbsp;</td>
  </tr>
 <tr>
    <td scope="row">&nbsp;</td>
  </tr>
  <tr>
    <td scope="row"><table width="100%" border="1" cellpadding="3" cellspacing="0">
      <tr>
        <td align="left" scope="row"><strong>Submitted By</strong></td>
        <td align="left" scope="row"><strong>Checked By:</strong></td>
        <td align="left" scope="row"><strong>Verified By:</strong></td>
        <td align="left" width="25%" scope="row"><strong>Approved By:</strong></td>
      </tr>
      <tr>
        <td height="100" align="left" scope="row"><%=tech_name%></td>
        
        <th align="left" scope="row">
        <%if  rpc_checkedby = "" or isnull(rpc_checkedby) then %>       
            <%if request.Cookies("GAPS")("slevel") = "sc" then  %>
              <input type=checkbox name="ClaimCheck_chk" value="OK">
              <input type="Submit" id="Checkby" name="Checkby" value="Click to Confirm" onclick="javascript:ApproveClaim('','action.asp?type=ClaimCheckOK&tech_code=<%=tech_code%>&rpc_id=<%=request("rpc_id")%>&techtype=TPC')">                                                                                                   
            <%end if %>
        <%else %>
              <%=rpc_checkedby %> <br/> 
              <%=rpc_checked_date %> 
        <%end if %>
        </th>
        
        <th align="left" scope="row"> 
        <%if rpc_verifiedby = "" or isnull(rpc_verifiedby) then %>    
          <%if Request.Cookies("GAPS")("verify_claim") = "Y" then %>
            <input type=checkbox name="ClaimVerify_chk" value="OK">
            <input type="Submit" id="Verifyby" name="Verifyby" value="Click to Confirm" onclick="javascript:ApproveClaim('','action.asp?type=ClaimVerifyOK&tech_code=<%=tech_code%>&rpc_id=<%=request("rpc_id")%>&techtype=TPC')">
           <%end if %>
         <%else %>
              <%=rpc_verifiedby %><br/>  
              <%=rpc_verified_date %> 
        <%end if %>
        </th>          
        
        <th align="left" scope="row">  
        </th>   
      </tr>         
        
    </table></td>
  </tr>
  <tr>
    <td scope="row">&nbsp;</td>
  </tr>
  <tr>
      <%if rpc_verifiedby <> "" and request.Cookies("GAPS")("slevel") <> "technician" then%>
           <td scope="row"><input type="submit" style="height:40px;width:200px;font-weight:bold" name="clink" id="clink" value="Copy Link" formaction='rm_rpt_tech_monthcommisionTPC_print.asp?rpc_id=<%=request("rpc_id")%>&jobyear=<%=jobyear%>&jobmonth=<%=jobmonth%>&tech_code=<%=tech_code%>&verified=<%=rpc_verifiedby%>&copylink=Yes' formmethod="post"/></td>;
      <%end if%>
  </tr>
  <tr>
    <td scope="row">&nbsp;</td>
  </tr>
</table>
</form>
</body>
</html>

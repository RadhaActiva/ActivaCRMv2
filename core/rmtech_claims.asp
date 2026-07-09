<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style4 {
            width: 67%;
        }
    </style>
</head>
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
    
if Request.Cookies("GAPS")("slevel") = "sc" then
    tech_code = request("tech_code")
 else
    tech_code = request.cookies("GAPS")("sloginid")
end if

submit_date = chkdate(date())

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

'this proc is needed to allow user to enter submit for 1st time and also to overwrite the record based on mm/yy if needed
set rs = server.CreateObject("adodb.recordset")

if Request.Cookies("GAPS")("slevel") = "sc" then
    sql1 = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_total_toll,tc_overwrty_amt,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2,tc_toll_receipt,tc_hotel_receipt,tc_parking_receipt,tc_fuel_receipt,tc_year_process,tc_month_process " &_
   " FROM tbltech_claim where tc_tech_code = '" & request("tech_code") & "' and tc_year = '" & request("jobyear") & "' and tc_month = '" & request("jobmonth") & "' and tc_claimID = '" & request("claim_id") &"'" 
else
    sql1 = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_total_toll,tc_overwrty_amt,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2,tc_toll_receipt,tc_hotel_receipt,tc_parking_receipt,tc_fuel_receipt,tc_year_process,tc_month_process " &_
    " FROM tbltech_claim where tc_tech_code = '" & request.cookies("GAPS")("sloginid") & "' and tc_year = '" & request("jobyear") & "' and tc_month = '" & request("jobmonth") & "' and tc_claimID = '" & request("claim_id") &"'" 
end if
    
rs.ActiveConnection = strconnect
rs.Source = sql1
rs.CursorLocation  = 3
rs.Open
if rs.eof then
    norecord = "There is no record found."
    stype = "addTechnClaim"
else    
    tech_code = rs("tc_tech_code")
    claim_id = rs("tc_claimID") 
    submit_date = rs("tc_submit_date")
    jobmonth = rs("tc_month")
    jobyear = rs("tc_year")
    jobmonth2 = rs("tc_month_process")
    jobyear2 = rs("tc_year_process")
    total_petrol = rs("tc_total_petrol")
    total_parking = rs("tc_total_parking")
    total_toll = rs("tc_total_toll")
    total_hotel = rs("tc_total_hotel") 
    desc1 = rs("tc_otherdesc1")
    descamt1 = rs("tc_otheramt1")
    desc2 = rs("tc_otherdesc2")
    descamt2 = rs("tc_otheramt2")
    deducdesc1 = rs("tc_deduc1")
    deducdescamt1 = rs("tc_deducamt1")
    deducdesc2 = rs("tc_deduc2")
    deducdescamt2 = rs("tc_deducamt2")
    total_extramileage = rs("tc_total_extramileage")
    overwrty_amt = rs("tc_overwrty_amt")
    toll_receipt = rs("tc_toll_receipt")
    hotel_receipt = rs("tc_hotel_receipt")
    parking_receipt = rs("tc_parking_receipt")
    fuel_receipt = rs("tc_fuel_receipt")
    jobmonth2 = rs("tc_month_process")
    jobyear2 = rs("tc_year_process")
end if
rs.Close
   
if request("claim_id") <> "" and request("tech_code") <> "" then        
    if Request.Cookies("GAPS")("slevel") = "sc" then
        sql1 = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_total_toll,tc_overwrty_amt,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2,tc_toll_receipt,tc_hotel_receipt,tc_parking_receipt,tc_fuel_receipt,tc_year_process,tc_month_process " &_
        " FROM tbltech_claim where tc_tech_code = '" & request("tech_code") &"' and tc_claimID = '" & request("claim_id") &"'" 
    else
        sql1 = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_total_toll,tc_overwrty_amt,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2,tc_toll_receipt,tc_hotel_receipt,tc_parking_receipt,tc_fuel_receipt,tc_year_process,tc_month_process " &_
        " FROM tbltech_claim where tc_tech_code = '" & request.cookies("GAPS")("sloginid") &"' and tc_claimID = '" & request("claim_id") &"'" 
    end if
        
        rs.ActiveConnection = strconnect
        rs.Source = sql1
        rs.CursorLocation  = 3
        rs.Open
        If Not rs.EOF Then
        claim_id = rs("tc_claimID") 
            submit_date = rs("tc_submit_date")
            jobmonth = rs("tc_month")
            jobyear = rs("tc_year")
            jobmonth2 = rs("tc_month_process")
            jobyear2 = rs("tc_year_process")
            total_petrol = rs("tc_total_petrol")
            total_parking = rs("tc_total_parking")
            total_toll = rs("tc_total_toll")
            total_hotel = rs("tc_total_hotel") 
            desc1 = rs("tc_otherdesc1")
            descamt1 = rs("tc_otheramt1")
            desc2 = rs("tc_otherdesc2")
            descamt2 = rs("tc_otheramt2")
            deducdesc1 = rs("tc_deduc1")
            deducdescamt1 = rs("tc_deducamt1")
            deducdesc2 = rs("tc_deduc2")
            deducdescamt2 = rs("tc_deducamt2")
            total_extramileage = rs("tc_total_extramileage")
            overwrty_amt = rs("tc_overwrty_amt")
            toll_receipt = rs("tc_toll_receipt")
            hotel_receipt = rs("tc_hotel_receipt")
            parking_receipt = rs("tc_parking_receipt")
            fuel_receipt = rs("tc_fuel_receipt")
            jobmonth2 = rs("tc_month_process")
            jobyear2 = rs("tc_year_process")
            rs.Close
	        'stype = "editPetrolClaim"	'button changes value if user's editing
	        'actionname = "Save" 
            'sbutton = "Update"            
        end if 
        'rs.close
end if

set rs = server.CreateObject("adodb.recordset")
if Request.Cookies("GAPS")("slevel") = "sc" then
    sql1 = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_total_toll,tc_overwrty_amt,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2,tc_toll_receipt,tc_hotel_receipt,tc_parking_receipt,tc_year_process,tc_month_process " &_
    " FROM tbltech_claim where tc_tech_code = '" & request("tech_code") &"'" 
else
    sql1 = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_total_toll,tc_overwrty_amt,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2,tc_toll_receipt,tc_hotel_receipt,tc_parking_receipt,tc_year_process,tc_month_process " &_
    " FROM tbltech_claim where tc_tech_code = '" & request.cookies("GAPS")("sloginid") &"'" 
end if 

rs.ActiveConnection = strconnect
rs.Source = sql1
rs.CursorLocation  = 3
rs.Open
If Not rs.EOF Then

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs.RecordCount
Remain = TotalRecord - Showed

If Remain > row Then
  LoopMax = Showed + row
Else
  LoopMax = Showed + Remain
End If

	If Int(TotalRecord/row) <> TotalRecord/row Then
	  pgCount = Int(TotalRecord/row) + 1
	Else
	  pgCount = TotalRecord/row
	End If

	if LoopMax mod row = 0 then
		pagestartno = LoopMax/row
	else
		pagestartno = pgCount
	end if		
end if

count = count + Showed
'link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby")
link = "&tech_code=" & request(tech_code) '& "&searchvalue=" & searchvalue & "&sortby=" & sortby & "&job_date_from=" & sp_date_from & "&job_date_to=" & sp_date_to & "&ordertype=" & ordertype

%> 

<script language="javascript">
    

    function RedirectURL_petrol()
    {
        claimid=document.getElementById('claim_id').value;
        if (claimid != "" )
        {
            window.open("rmtech_petrol_claims.asp?reset", '_blank', "left=100,width=900,height=600");
            return false;
        }
    }

    function RedirectURL_hotel()
    {
        claimid = document.getElementById('claim_id').value;
        if (claimid != "") {
            window.open("rmtech_hotel_claims.asp?reset", '_blank', "left=100,width=900,height=600");
            return false;
        }
    }

    function RedirectURL_parkingtoll()
    {
        claimid = document.getElementById('claim_id').value;
        if (claimid != "") {
            window.open("rmtech_parkingtoll_claims.asp?reset", '_blank', "left=100,width=900,height=600");
            return false;
        }     
    }

    function RedirectURL_exmileage()
    {
        claimid = document.getElementById('claim_id').value;
        if (claimid != "") {
            window.open("rmtech_exmileage_claims.asp?reset", '_blank', "left=100,width=900,height=600");
            return false;
        }
    }
</script>
        <tr>
          <td><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left">Submit Claim Request</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                    
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="action.asp?type=<%=stype%>&tech_code=<%=tech_code%>" enctype="multipart/form-data">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                       <tr>
                        <td width="20%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Claim ID</font></div></td> &nbsp;&nbsp;&nbsp
                        <td class="auto-style4"><input name="claim_id" type="text" id="claim_id" size="15" maxlength="7" style="background-color: #cccccc;" value="<%=claim_id%>" readonly /></td>
                      </tr>
                       <tr>
                         <td width="20%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Technician Code</font></div></td> &nbsp;&nbsp;&nbsp
                        <td class="auto-style4"><input name="tech_code" type="text" id="tech_code" style="background-color: #cccccc;" value="<%=tech_code%>" size="15" readonly/></td>
                      </tr>
                      <tr>
                        <td width="20%" bgcolor="#CD6155"><font color="#FFFFFF"> Claim Date</font><br/></td> &nbsp;&nbsp;&nbsp
                        <td class="auto-style4"><align="left"><color="#000000">
                          <input name="submit_date" type="text" id="submit_date" style="background-color: #cccccc;" value="<%=submit_date%>" size="15" readonly/>
                        </tr>
                      <tr>
                          <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                      </tr>

                      <tr>
                      <td width="20%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Claim Submission Month</font></div></td> &nbsp;&nbsp;&nbsp      
                      <td class="auto-style4"><align="left"><color="#000000">
                        <select name="jobyear" id="jobyear">
                        <option value="2022"<%if jobyear="2022" then response.write " selected"%>>2022</option>
                        <option value="2023"<%if jobyear="2023" then response.write " selected"%>>2023</option>
                        <option value="2024"<%if jobyear="2024" then response.write " selected"%>>2024</option>
                        <option value="2025"<%if jobyear="2025" then response.write " selected"%>>2025</option>
                        <option value="2026"<%if jobyear="2026" then response.write " selected"%>>2026</option>
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
                         <%if Request.Cookies("GAPS")("slevel") = "sc" then %>
                                &nbsp; &nbsp;<Strong>Claim Processing Month </Strong>
                             <select name="jobyear2" id="jobyear2">
                                <option value="2022"<%if jobyear2="2022" then response.write " selected"%>>2022</option>
                                <option value="2023"<%if jobyear2="2023" then response.write " selected"%>>2023</option>
                                <option value="2024"<%if jobyear2="2024" then response.write " selected"%>>2024</option>
                                <option value="2025"<%if jobyear2="2025" then response.write " selected"%>>2025</option>
                                <option value="2026"<%if jobyear2="2026" then response.write " selected"%>>2026</option>
                            </select>
                            <select name="jobmonth2" id="jobmonth2">
                              <option value="1" <%if jobmonth2="1" then response.write " selected"%>>Jan</option>
                              <option value="2" <%if jobmonth2="2" then response.write " selected"%>>Feb</option>
                              <option value="3" <%if jobmonth2="3" then response.write " selected"%>>Mar</option>
                              <option value="4" <%if jobmonth2="4" then response.write " selected"%>>Apr</option>
                              <option value="5" <%if jobmonth2="5" then response.write " selected"%>>May</option>
                              <option value="6" <%if jobmonth2="6" then response.write " selected"%>>Jun</option>
                              <option value="7" <%if jobmonth2="7" then response.write " selected"%>>Jul</option>
                              <option value="8" <%if jobmonth2="8" then response.write " selected"%>>Aug</option>
                              <option value="9" <%if jobmonth2="9" then response.write " selected"%>>Sep</option>
                              <option value="10" <%if jobmonth2="10" then response.write " selected"%>>Oct</option>
                              <option value="11" <%if jobmonth2="11" then response.write " selected"%>>Nov</option>
                              <option value="12" <%if jobmonth2="12" then response.write " selected"%>>Dec</option>
                            </select>
                         <%end if%>
                        </td>
                      </tr>

                      <tr>
                          <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                      </tr>

                      <tr>
                        <td  bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Total Petrol </font></div></td> &nbsp;&nbsp
                        <td class="auto-style4"><input name="total_petrol" type="text" id="total_petrol" value="<%=ChkNumber2(total_petrol)%>" maxlength="8" size="8" /> (RM) (Approved Amount is 90%)</td>
                      </tr>
                       <tr>
                       <td bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Total Parking </font></div></td> &nbsp;&nbsp
                        <td class="auto-style4"><input name="total_parking" type="text" id="total_parking" value="<%=ChkNumber2(total_parking)%>" maxlength="8" size="8" /> (RM)
                        </td>
                      </tr>
                        <tr>
                        <td bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Toll Amount</font></div></td> &nbsp;&nbsp;&nbsp
                        <td class="auto-style4"><input name="total_toll" type="text" id="total_toll" value="<%=ChkNumber2(total_toll)%>" maxlength="8" size="8" /> (RM)
                        </td>
                      </tr>
                        <tr>
                        <td bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Hotel Expenses</font></div></td> &nbsp;&nbsp;&nbsp
                        <td class="auto-style4"><input name="total_hotel" type="text" id="total_hotel" value="<%=ChkNumber2(total_hotel)%>" maxlength="8" size="8" /> (RM)
                        </td>
                        </tr>
                        <tr>
                        <td bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Extra Mileage Service </font> </div></td> &nbsp;&nbsp;&nbsp
                        <td class="auto-style4"><input name="total_extramileage" type="text" id="total_extramileage" value="<%=ChkNumber2(total_extramileage)%>" size="8" /> (KM)</td>
                        </tr>
                        <tr>
                         <td colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                       <tr>
                        <td class="titleblue1" colspan="2"><div align="left">Other Additional Claims</div></td>                    
                      </tr>
                         <tr>
                         <td colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                         <tr>
                       <td bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Desc 1</font></div></td>
                        <td class="auto-style4"><textarea rows = "2" cols = "100" name="desc1"><%=desc1%></textarea></td>        
          
                        <td><input name="descamt1" type="text" id="descamt1" value="<%=descamt1%>" maxlength="8" size="6" />&nbsp;&nbsp(RM)</td>
                      </tr>
                        <tr>
                        <td bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Desc 2</font></div></td>
                        <td class="auto-style4"> <textarea rows = "2" cols = "100" name="desc2" ><%=desc2%></textarea></td>
                        <td><input name="descamt2" type="text" id="descamt2" value="<%=descamt2%>" maxlength="8" size="6" />&nbsp;&nbsp(RM)</td>
                      </tr>
                          <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                        </table>
                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                        <td width="50%" class="titleblue1"><div align="left">Over warrranty overriding fees</div></td>                    
                        </tr>
                      </table>
                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                         <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                         <tr>
                        <td width="20%" bgcolor="#CD6155"><font color="#FFFFFF">Over-Warranty Service Case & Spare-Part</font></td>
                        <td width="50%"><input name="overwrty_amt" type="text" id="overwrty_amt" value="<%=ChkNumber2(overwrty_amt)%>" maxlength="8" size="8" />&nbsp;&nbsp;(RM)</td>
                      </tr>
                        <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                       <tr>
                        <td class="titleblue1""><div align="left">Other Deductions</div></td>                    
                      </tr>
                         <tr>
                         <td bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                         <tr>
                        <td bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Description 1 </font></div></td>
                        <td><input name="deducdesc1" type="text" id="deducdesc1" value="<%=deducdesc1%>" size="80" /></td><td width="3%"></td>
                        <td><input name="deducdescamt1" type="text" id="deducdescamt1" value="<%=deducdescamt1%>" maxlength="8" size="6" />&nbsp;&nbsp(RM)</td>
                      </tr>
                        <tr>
                        <td bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Description 2 </font></div></td>
                        <td><input name="deducdesc2" type="text" id="deducdesc2" value="<%=deducdesc2%>" size="80" /></td><td width="3%"></td>
                        <td><input name="deducdescamt2" type="text" id="deducdescamt2" value="<%=deducdescamt2%>" maxlength="8" size="6" />&nbsp;&nbsp(RM)</td>
                      </tr>   
                        <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                       <tr>                     
                         <tr>
                        <td class="titleblue1"><div align="left">Upload Receipts</div>&nbsp;&nbsp</td><td><div align="left">(Upload and Submit the images one-by-one if you have multiple receipts)</div></td>                    
                      </tr>
                         <tr>
                         <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>                         
                        </tr>
                        <tr>
                        <td><div align="left">Attach Toll Receipt</div></td>
                        <td><input type="file" name="toll_receipt" id="toll_receipt" />&nbsp;&nbsp<%=toll_receipt%></td>
                        </tr>
                        <tr>
                        <td><div align="left">Attach Hotel Receipt</div></td>
                        <td><input type="file" name="hotel_receipt" id="hotel_receipt" />&nbsp;&nbsp<%=hotel_receipt%></td>
                        </tr>
                        <tr>
                        <td><div align="left">Attach Parking Receipt</div></td>
                        <td><input type="file" name="parking_receipt" id="parking_receipt" />&nbsp;&nbsp<%=parking_receipt%></td>
                        </tr>
                            <tr>
                        <td><div align="left">Attach Fuel Receipt</div></td>
                        <td><input type="file" name="fuel_receipt" id="fuel_receipt" />&nbsp;&nbsp<%=fuel_receipt%></td>
                        </tr>
                         <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                         <tr>                         
                         <td colspan="2" valign="top" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="Submit"></td>
                        </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">
                            <table width="100%" border="0" cellspacing="0">
                             <tr> <td class="titleblue1"><div align="left"><strong>Additional Claims</strong></div></td>
                             <td>
                             <input style="height:40px;width:150px;font-weight:bold;background-color: #1b6bcf; color: white;" type="submit" name="Submit" onclick="javascript:return RedirectURL_petrol();" value="Petrol">
                             <input style="height:40px;width:150px;font-weight:bold;background-color: #1b6bcf; color: white;" type="submit" name="Submit" onclick="javascript:return RedirectURL_hotel();" value="Hotel">
                             <input style="height:40px;width:150px;font-weight:bold;background-color: #1b6bcf; color: white;" type="submit" name="Submit" onclick="javascript:return RedirectURL_parkingtoll();" value="Parking/Toll">
                             <input style="height:40px;width:150px;font-weight:bold;background-color: #1b6bcf; color: white;" type="submit" name="Submit" onclick="javascript:return RedirectURL_exmileage();" value="Ex Mileage">
                             </td>
                             </tr>
                         </table>
                  </td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rmtech_claims.asp?num=" & (j-1) * row & link & "&sp_status=" & sp_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rmtech_claims.asp?num=" & Showed+row & link & "&sp_status=" & sp_status & "'> Next >></a>"
	End If%>

                  </td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table border="0" align="right" cellpadding="5" cellspacing="1">
                    <tr>
                      
                    </tr>
                  </table>
                   </td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" width="5%" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Claim ID</font></td>
                      <td align="center" width="10%" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Submission Date</font></td>
                      <td align="center" width="10%" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Submission MM/YY</font></td>
                      <td align="center" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Total Petrol</font></td>
                      <td align="center" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Total Parking</font></td>
                      <td align="center" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Toll Amount<br/></font></td>
                      <td align="center" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Hotel Expenses</font></td>
                      <td align="center" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Extra Mileage Service</font></td>
                      <td align="center" bgcolor="#475387 " width="10%" class="style1"><font color="#FFFFFF">Other Claims Amt 1</font></td>
                      <td align="center" width="10%" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Other Claims Amt 2</font></td>
                      <td align="center" width="10%" bgcolor="#475387 " class="style1"><font color="#FFFFFF">O/Wrty<br/> Claims</font></td>
                      <td align="center" width="10%" bgcolor="#475387 " class="style1"><font color="#FFFFFF">Other Deduction Amt 1</font></td>
                      <td align="center" width="10%" bgcolor="#475387" class="style1"><font color="#FFFFFF">Other Deduction Amt 2</font></td>
                      <td align="left" width="5%" bgcolor="#475387" class="style1"><font color="#FFFFFF">Action</font></td>
                    </tr>                    
<% 
job_totalAmt = 0
if not rs.eof then
rs.Move Showed
count = Showed + 1
end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
                    <tr bgcolor="#FFFFFF">
                      <td align="left"><%=rs("tc_claimID")%></td>
                      <td> <%=chkdate(rs("tc_submit_date"))%></td>
                      <td> <%=rs("tc_month")%>/<%=rs("tc_year")%></td>
                      <td> <%=rs("tc_total_petrol")%></td>
                      <td> <%=rs("tc_total_parking")%></td>
                      <td> <%=chknumber2(rs("tc_total_toll"))%></td>
                      <td> <%=chknumber2(rs("tc_total_hotel"))%> </td>
                      <td> <%=chknumber2(rs("tc_total_extramileage"))%> </td>
                      <td align="center"> <%=chknumber2(rs("tc_otheramt1"))%></td>
                      <td align="center" nowrap="nowrap"><%=chknumber2(rs("tc_otheramt2"))%></td>
                      <td align="center" nowrap="nowrap"><%=chknumber2(rs("tc_overwrty_amt"))%></td>
                      <td align="right"> <%=chknumber2(rs("tc_deducamt1"))%> </td>
                      <td align="right"> <%=chknumber2(rs("tc_deducamt2"))%> </td>                      
                      <td align="left"> <input type="button" name="aditem211" id="aditem211" value="Edit" onclick="document.location.href='rmtech_claims.asp?claim_id=<%=rs("tc_claimID")%>&tech_code=<%=tech_code%>#claims'"/></td>
                    </tr>
                  
<%
'sp_totalqty = sp_totalqty + rs("sp_totalqty")
'sp_totalAmt = sp_totalAmt + rs("sp_totalAmt")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>

  					<tr>
                    
                  </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rmtech_claims.asp?num=" & (j-1) * row & link & "&sp_status=" & sp_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rmtech_claims.asp?num=" & Showed+row & link & "&sp_status=" & sp_status & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
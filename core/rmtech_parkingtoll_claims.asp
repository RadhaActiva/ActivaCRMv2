<!-- #include file="database/datastore.asp" -->
<%

'trip_date = request("trip_date")
'mileage_start  = request("mileage_start")
'mileage_end = request("mileage_end")
'distance = request("distance")
'vehicle_no = request("vehicle_no")
parking_amount  = request("parking_amount")
parking_amount  = request("toll_amount")
claim_id  = request("claim_no")

tech_code = request.cookies("GAPS")("sloginid")
submit_date = chkdate(date())

'if request("total_mileage") <> "" or request("total_toll") <> "" or request("total_hotel") <> "" or request("descamt1") <> "" or request("descamt2") <> "" or request("deducdescamt1") <> "" or request("deducdescamt2") <> "" then	  
'if request("total_mileage") <> "" or request("total_toll") <> "" or request("total_hotel") <> "" then
'    stype = "addTechnClaim"
'   actionname = "Save" 
'end if
    
'if request.Cookies("GAPS")("slevel") = "technician" then 
 '  sql = sql & " and sp_tech_code =   "
'end if

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

'default value of the button
 stype = "addParkingTollClaim"	
 actionname = "Save" 
 sbutton = "Add"

set rs = server.CreateObject("adodb.recordset")

if request("claim_id") <> "" and request("tech_code") <> "" then
        sql1 = "SELECT tpt_id,tpt_date,tpt_year,tpt_month,tpt_tech_code,tpt_job_sheet,tpt_parking_amount, tpt_toll_amount,tpt_trip_date " & _      
            "FROM tbltech_claim_parkingtoll where tpt_tech_code = '" & request.cookies("GAPS")("sloginid") & "' and tpt_id = '" & request("claim_id") & "'"
        rs.ActiveConnection = strconnect
        rs.Source = sql1
        rs.CursorLocation  = 3
        rs.Open
        If Not rs.EOF Then
	        claim_id = rs("tpt_id") 
            jobmonth = rs("tpt_month")
            jobyear = rs("tpt_year")
            trip_date = rs("tpt_trip_date")
            submit_date = rs("tpt_date")
            job_sheet = rs("tpt_job_sheet")
            tech_code= rs("tpt_tech_code")
            parking_amount = rs("tpt_parking_amount")
            toll_amount = rs("tpt_toll_amount")
            rs.Close
	        stype = "editParkingTollClaim"	'button changes value if user's editing
	        actionname = "Save" 
            sbutton = "Update"
        end if
end if

sql1 = "SELECT tpt_id,tpt_date,tpt_year,tpt_month,tpt_tech_code,tpt_job_sheet,tpt_parking_amount, tpt_toll_amount,tpt_trip_date " & _      
       "FROM tbltech_claim_parkingtoll where tpt_tech_code = '" & request.cookies("GAPS")("sloginid") & "' order by tpt_date"
rs.ActiveConnection = strconnect
rs.Source = sql1
rs.CursorLocation  = 3
rs.Open
        'If Not rs.EOF Then
	     '   claim_id = rs("tp_id") 
         '   jobmonth = rs("tp_month")
          '  jobyear = rs("tp_year")
          '  vehicle_no = rs("tp_vehicle_no")
          '  mileage_start = rs("tp_mileage_start")
          '  mileage_end = rs("tp_mileage_end")
          '  distance = rs("tp_distance")
          '  tech_node= rs("tp_tech_code")
          '  submit_date = rs("tp_date")
           ' trip_date= rs("tp_trip_date")
        'end if


If Not rs.EOF Then
if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 10
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

function isEmpty(s) {
        return ((s == null) || (s.length == 0));
 }
function CalcMileage(id,orderlinks,otype) 
{
    //document.forms["form1"].distance.value = document.forms["form1"].mileage_end.value - document.forms["form1"].mileage_start.value
    document.form1.action = orderlinks;
    document.form1.submit()
}

</script>


        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td><div align="left"><font color="#CC0000"><strong>Parking and Toll Claim Form</strong></div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                    
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="action.asp?type=<%=stype%>" >
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                       <tr>
                        <td width="10%" bgcolor="#CD6155"> <div align="left"><font color="#FFFFFF">Technician Code</font></div></td>
                        <td><input name="tech_code" type="text" id="tech_code" style="background-color: #cccccc;" value="<%=tech_code%>" size="15" readonly/></td>
                           <td width="20%" align="left" class="titlegrey1">
                            <a href="rmtech_parkingtoll_claims_excel.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&tech_code=<%=tech_code%>" target="_blank">
                            <img src="images/excel.jpg" width="57" height="21" border="0" />
                      </tr>
                      <tr>
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF">Submission Date</font><br/></td>
                        <td width="40%"><align="left">
                            <input name="submit_date" type="text" id="submit_date" style="background-color: #cccccc;" value="<%=chkdate(submit_date)%>" size="15" readonly />
                             &nbsp;&nbsp;
                    Claim Submission Month 
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
                      </td>
                      </tr>
                         <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                       <tr>
                        <td width="16%" class="titleblue1"><div align="left">Trip Details</div></td>                    
                      </tr>
                         <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                         <tr>
                        <td width="16%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Trans ID</font></div></td>
                        <td><input name="claim_id" type="text" id="claim_id" size="12" maxlength="7" style="background-color: #cccccc;" value="<%=claim_id%>" readonly /></td><td width="5%"></td>
                      </tr>
                         <tr>
                        <td width="10%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Trip Date</font></div></td>
                        <td><input name="trip_date" type="text" id="trip_date" size="12" readonly value="<%=chkdate(trip_date)%>" />
                            <a href="javascript:void(null)" onclick="window.dateField = document.form1.trip_date;calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>
                       </td>
                      </tr>
                       <tr>
                        <td width="10%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Job Sheet No</font></div></td>
                        <td><select name="job_sheet" id="job_sheet"><option value="0"></option>
                        <%  
                        sql2 = "Select job_code FROM tbljob where job_submitforclaims='Yes' and job_tech_code='" & tech_code & "' and job_claim_approved is NULL"	
                        set rs1 = server.CreateObject("adodb.recordset")
				        rs1.Open sql2,strconnect,3,3,&H0001
                        while Not rs1.EOF
					          if (job_sheet) = rs1("job_code") then
					          response.write "<option value='" & rs1("job_code") & "' selected>" & rs1("job_code") & "</option>"
					          else
					          response.write "<option value='" & rs1("job_code") & "'>" & rs1("job_code") & "</option>"
					          end if 					  
				        rs1.movenext
				        wend
				        rs1.close	
			    	    %>
                        </select>
                       </td>
                      </tr>
                        <tr>
                        <td width="16%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Parking Amount</font></div></td>
                        <td><input name="parking_amount" type="text" id="parking_amount" size="12" maxlength="7" value="<%=parking_amount%>" /></td><td width="5%"></td>
                      </tr>
                        <tr>
                        <td width="16%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Toll Amount</font></div></td>
                        <td><input name="toll_amount" type="text" id="toll_amount" size="12" maxlength="7" value="<%=toll_amount%>" /></td><td width="5%"></td>
                      </tr>
                         <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                         <tr>                          
                         <td colspan="2" valign="top" bgcolor="#FFFFFF"><input type="submit" name="AddParkingToll" value="<%=sbutton%>" onclick="javascript:CalcMileage('','action.asp?type=<%=stype%>','')"/> </td>
                        </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rmtech_parkingtoll_claims.asp?num=" & (j-1) * row & link & "&sp_status=" & sp_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rmtech_parkingtoll_claims.asp?num=" & Showed+row & link & "&sp_status=" & sp_status & "'> Next >></a>"
	End If%>

                  </td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table border="1" align="right" cellpadding="5" cellspacing="1">
                    <tr>
                      
                    </tr>
                  </table>
                   </td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="left" width="10%" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Trans ID</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Submission Date</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>MM / YY</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Trip Date</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Job Sheet</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Parking Amount</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Toll Amount</strong></font></td>
                      <td align="left" width="10%" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Action</strong></font></td>
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
                      <td align="left"><%=rs("tpt_id")%></td>
                      <td> <%=chkdate(rs("tpt_date"))%></td>
                      <td> <%=rs("tpt_month")%>/<%=rs("tpt_year")%></td>
                      <td> <%=chkdate(rs("tpt_trip_date"))%></td>
                      <td> <%=rs("tpt_job_sheet")%> </td>
                      <td> <%=ChkNumber2(rs("tpt_parking_amount"))%> </td>
                      <td> <%=ChkNumber2(rs("tpt_toll_amount"))%> </td>
                      <td>
                          <input type="button" name="aditem211" id="aditem211" value="Edit" onclick="document.location.href='rmtech_parkingtoll_claims.asp?claim_id=<%=rs("tpt_id")%>&tech_code=<%=tech_code%>#claims'"/>
                          <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmDel('<%=rs("tpt_id")%>','action.asp?type=delParkingTollClaim&claim_id=<%=rs("tpt_id")%>&tech_code=<%=tech_code%>&jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>')"/>                                                                                 
                      </td>
                    </tr>
                  
<%

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
					Response.Write " <a href='rmtech_parkingtoll_claims.asp?num=" & (j-1) * row & link & "&sp_status=" & sp_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rmtech_parkingtoll_claims.asp?num=" & Showed+row & link & "&sp_status=" & sp_status & "'> Next >></a>"
	End If%>

              </td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
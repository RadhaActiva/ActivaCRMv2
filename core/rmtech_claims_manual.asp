<!-- #include file="header.asp" -->
<!-- #include file="database/datastore.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            width: 142px;
        }
        .auto-style2 {
            width: 142px;
            height: 13px;
        }
        .auto-style3 {
            height: 22px;
        }
    </style>
</head>
<%

claims_id  = request("claims_id")
tech_code = request("tech_code")
trip_date = request("trip_date")
total_petrol  = request("total_petrol")
total_toll = request("total_toll")
total_parking = request("total_parking")
total_incentive = request("total_incentive")
completed  = request("completed")
period  = request("period")
entry_date  = request("entry_date")

'tech_code = request.cookies("GAPS")("sloginid")
submit_date = chkdate(date())

'default value of the button
 stype = "addManualClaim"	
 actionname = "Save" 
 sbutton = "Add"

set rs = server.CreateObject("adodb.recordset")

if request("claims_id") <> "" then
        sql1 = "SELECT claims_id, tech_code, total_petrol,total_toll, total_parking,total_incentive, completed, period, entry_date " & _      
            "FROM tbltech_claim_manual where claims_id = '" & request("claims_id") & "'"
        rs.ActiveConnection = strconnect
        rs.Source = sql1
        rs.CursorLocation  = 3
        rs.Open
        If Not rs.EOF Then
            claims_id  = rs("claims_id")
            tech_code = rs("tech_code")
            total_petrol  = rs("total_petrol")
            total_toll = rs("total_toll")
            total_parking = rs("total_parking")
            total_incentive = rs("total_incentive")
            completed  = rs("completed")
            period  = rs("period")
            entry_date  = rs("entry_date")
            rs.Close
	        stype = "editManualClaim"	'button changes value if user's editing
	        actionname = "Save" 
            sbutton = "Update"
        end if
end if

sql1 = "SELECT claims_id, tech_code, total_petrol,total_toll, total_parking,total_incentive, completed, period, entry_date " & _      
       "FROM tbltech_claim_manual order by claims_id desc"
rs.ActiveConnection = strconnect
rs.Source = sql1
rs.CursorLocation  = 3
rs.Open

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
link = "&tech_code=" & request(tech_code) '& "&searchvalue=" & searchvalue & "&sortby=" & sortby & "&job_date_from=" & sp_date_from & "&job_date_to=" & sp_date_to & "&ordertype=" & ordertype
%> 

<script language="javascript">

function isEmpty(s) {
        return ((s == null) || (s.length == 0));
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Technicians Manual Claims Entry</div></td>
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
                        <td width="30%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Trans ID</font></div></td>
                        <td><input name="claim_id" type="text" id="claim_id" size="12" maxlength="7" style="background-color: #cccccc;" value="<%=claims_id%>" readonly /></td><td width="5%"></td>
                      </tr>
                       <tr>
                        <td width="20%" bgcolor="#CD6155"> <div align="left"><font color="#FFFFFF">Technician Code</font></div></td>
                        <td><select name="tech_code" id="tech_code">
                           <option value="">--Select Technician--</option>
                            <%			
			    sql3 = "select distinct tbljob.job_tech_code, tbltechnician.tech_name FROM tbljob INNER JOIN " & _
                "tbltechnician ON tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null and tbljob.job_status='Posted' " & _
                "and tbljob.job_submitforclaims='Yes' and tbltechnician.tech_type = 'IHT'"
                set rs3 = server.CreateObject("adodb.recordset")
				rs3.Open sql3,strconnect,3,3,&H0001
                while Not rs3.EOF
					  if rs3("job_tech_code") = job_tech_code then
					  response.write "<option value='" & rs3("job_tech_code") & "' selected>" & rs3("job_tech_code") & " - " & rs3("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs3("job_tech_code") & "'>" & rs3("job_tech_code") & " - " & rs3("tech_name")  & "</option>"
					  end if 					  
				rs3.movenext
				wend
				rs3.close					
				%>      </select> </td>
                      </tr>
                      <tr>
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF">Period Used</font><br/></td>
                        <td width="40%"><align="left">
                            <input name="submit_date" type="text" id="submit_date" style="background-color: #cccccc;" value="<%=period%>" size="15" readonly />
                      </td>
                      </tr>
                         <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                       <tr>
                        <td width="16%" class="titleblue1"><div align="left">Claims Amount</div></td>                    
                      </tr>
                         <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                       
                         <tr>
                        <td width="30%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Total Petrol Claims</font></div></td>
                        <td><input name="total_petrol" type="text" id="total_petrol" size="12" value="<%=total_petrol%>" /></td>
                      </tr>
                        <tr>
                        <td width="30%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Total Toll Claims</font></div></td>
                        <td><input name="total_toll" type="text" id="total_toll" size="12" maxlength="7" value="<%=total_toll%>" /></td>
                      </tr>
                        <tr>
                        <td width="16%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Total Parking Claim</font></div></td>
                        <td><input name="total_parking" type="text" id="total_parking" size="12" maxlength="12" value="<%=total_parking%>" /></td>
                      </tr>
                        <tr>
                        <td width="16%" bgcolor="#CD6155" class="auto-style3"><div align="left"><font color="#FFFFFF">Total Target Incentive Amount</font></div></td>
                        <td class="auto-style3"><input name="total_incentive" type="text" id="total_incentive" size="12" maxlength="12" value="<%=total_incentive%>" /></td>
                      </tr>
                        <tr>
                        <td width="16%" bgcolor="#CD6155" class="auto-style3"><div align="left"><font color="#FFFFFF">Entry Date</font></div></td>
                        <td class="auto-style3"><input name="entry_date" type="text" id="entry_date" size="12" readonly style="background-color: #cccccc;" value="<%=entry_date%>" class="auto-style1" /></td>
                      </tr>      
                        <tr>
                        <td width="16%" bgcolor="#CD6155"><div align="left"><font color="#FFFFFF">Used in Claims</font></div></td>
                        <td><input name="completed" type="text" id="completed" size="12" readonly style="background-color: #cccccc;" value="<%=completed%>" class="auto-style2"/></td>
                      </tr>      
                         <tr>
                         <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                         <tr>                          
                         <td colspan="2" valign="top" bgcolor="#FFFFFF"><input type="submit" name="AddClaims" value="<%=sbutton%>"/> </td>
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
					Response.Write " <a href='rmtech_petrol_claims.asp?num=" & (j-1) * row & link & "&sp_status=" & sp_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rmtech_petrol_claims.asp?num=" & Showed+row & link & "&sp_status=" & sp_status & "'> Next >></a>"
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
                      <td align="left" width="10%" bgcolor="#475387" class="style1"><font color="#FFFFFF">Claims ID</font></td>
                      <td align="left" width="10%" bgcolor="#475387" class="style1"><font color="#FFFFFF">Tech Code</font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Entry Date</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Period Used</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Petrol Amt</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Toll Amt</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Parking Amt</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Target Incentive Amt</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Used</strong></font></td>
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
                      <td align="left"><%=rs("claims_id")%></td>
                      <td align="left"><%=rs("tech_code")%></td>
                      <td> <%=chkdate(rs("entry_date"))%></td>
                      <td> <%=rs("period")%></td>
                      <td> <%=ChkNumber2Decimal(rs("total_petrol"))%></td>
                      <td> <%=ChkNumber2Decimal(rs("total_toll"))%> </td>
                      <td> <%=ChkNumber2Decimal(rs("total_parking"))%> </td>
                      <td> <%=ChkNumber2Decimal(rs("total_incentive"))%> </td>

                        <td align="center" <% If rs("completed") = "No" Then %> style="color:#CC0000;" <% End If %>>
                            <%= rs("completed") %>
                        </td>

                      <td>
                          <input type="button" name="aditem211" id="aditem211" value="Edit" onclick="document.location.href='rmtech_claims_manual.asp?claims_id=<%=rs("claims_id")%>#claims'"/>
                          <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmDel('<%=rs("claims_id")%>','action.asp?type=delClaimsManual&claim_id=<%=rs("claims_id")%>&tech_code=<%=tech_code%>')"/>                                                                                 
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
					Response.Write " <a href='rmtech_claims_manual.asp?num=" & (j-1) * row & link & "&sp_status=" & sp_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rmtech_claims_manual.asp?num=" & Showed+row & link & "&sp_status=" & sp_status & "'> Next >></a>"
	End If%>

              </td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
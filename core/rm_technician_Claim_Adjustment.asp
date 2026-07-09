<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            width: 642px;
        }
        .auto-style2 {
            width: 119px;
        }
    </style>
</head>
<%
job_tech_type = request("job_tech_type")
tech_type = request("tech_type")
job_tech_code = request("job_tech_code")


if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

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

i = 1	


 sql2 = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_total_toll,tc_overwrty_amt,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2,tc_toll_receipt,tc_hotel_receipt,tc_parking_receipt,tc_year_process,tc_month_process " &_
 "FROM tbltech_claim"

if job_tech_code <> "All" and job_tech_code <> "" then
  sql2 = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_total_toll,tc_overwrty_amt,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2,tc_toll_receipt,tc_hotel_receipt,tc_parking_receipt,tc_year_process,tc_month_process " &_
    "FROM tbltech_claim where tc_tech_code = '" & job_tech_code & "'"
end if
 
set rs4 = server.CreateObject("adodb.recordset")
rs4.ActiveConnection = strconnect
rs4.Source = sql2
rs4.CursorLocation  = 3
rs4.Open
if rs4.eof then
   norecord = "There is no record found."
end if

If Not rs4.EOF Then

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs4.RecordCount
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

link = "&job_tech_code=" & job_tech_code & " &job_tech_code=" & job_tech_code & "&tech_type=" & tech_type &"&job_date_from=" & job_date_from & "&job_date_to =" & job_date_to & "&jobyear=" & jobyear & "&jobmonth=" & jobmonth

%>  
        <tr>
          <td><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Technician </font>Claim Adjustment Screen</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td class="titleblue1">&nbsp;</td>
                      <td width="20%" align="center" class="titlegrey1"></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_technician_Claim_Adjustment.asp?type=reset">
                    <table width="90%" border="0" cellpadding="0" cellspacing="0">
                        <td width="16%" class="titlegrey1">Technician</td>
                        <td class="auto-style2"><span class="titlegrey1">
                          <select name="job_tech_code" id="job_tech_code">
                           <option value="All">All</option>
                            <%			
				'sql3 = "SELECT tech_code, tech_name FROM tbltechnician where tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC'"
                sql3 = "select distinct tbljob.job_tech_code, tbltechnician.tech_name FROM tbljob INNER JOIN " & _
                "tbltechnician ON tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null and tbljob.job_status='Posted' " & _
                "and tbljob.job_submitforclaims='Yes'"
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
				%>      </select>
                        </span></td> &nbsp;&nbsp;&nbsp;                     
                        <td><span class="titlegrey1">
                          <input type="submit" name="button2" id="button3" value="Display Claims" />
                        </span></td>
                      </tr>                    
                    </table>
                  </form></td>
                </tr>               
                <tr>
                  <td align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_technician_Claim_Adjustment.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_technician_Claim_Adjustment.asp.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If      
 %>

                </td>
                </tr>
                <tr>    
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table border="0" cellpadding="4" cellspacing="0" class="auto-style1">
                    <tr>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span>No</span></strong></font></td>
                      <td align="center" width="30%" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span>Claim ID</span></strong></font></td>
                      <td align="left" nowrap="nowrap"    bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span> Tech Code</span></strong></font></td>
                      <td align="center" width="60%" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span>Tech Name</span></strong></font></td>
                      <td align="center" width="30%" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span>Submission <br/>Month</span></strong></font></td>
                      <td align="center" width="30%" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Submission <br/>Year</strong></font></td>                      
                      <td align="center" width="30%" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span>Processing <br/>Month</span></strong></font></td>
                      <td align="center" width="30%" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Processing <br/>Year</strong></font></td>     
                    </tr>   
<%

if not rs4.eof then
rs4.Move Showed
count = Showed + 1
end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

     techname = ""
     sql = "select tech_name from tbltechnician where tech_code = '" & rs4("tc_tech_code") & "'  "
     techname = selectid(sql)
%>                    
                    <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center" valign="top" nowrap="nowrap"><%=j%></td>
                      <td align="center" valign="top" nowrap="nowrap"><%=rs4("tc_claimID")%></td>
                      <td align="center" valign="top" nowrap="nowrap"><strong><font color="#0000FF"><a href="tech_claims.asp?tech_code=<%=rs4("tc_tech_code")%>&jobmonth=<%=rs4("tc_month")%>&jobyear=<%=rs4("tc_year")%>" target="_blank"><%=rs4("tc_tech_code")%></a></font></strong></td>
                      <td align="center"  valign="top" nowrap="nowrap"><%=techname%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs4("tc_month")%></td>
                      <td align="center" " valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs4("tc_year")%></td>					                          
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs4("tc_month_process")%></td>
                      <td align="center" " valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs4("tc_year_process")%></td>	
                    </tr>      
 <%
count = count + 1 
i = i + 1
rs4.MoveNext
Next
rs4.Close
%> 
<tr bgcolor="<%=nbgcolor%>">                    
</tr>                
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
					Response.Write " <a href='rm_technician_Claim_Adjustment.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_technician_Claim_Adjustment.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->

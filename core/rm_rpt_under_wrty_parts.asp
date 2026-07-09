<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            color: #FFFFFF;
        }
    </style>
</head>
<%
orderby = request("orderby")
ordertype = request("ordertype")
cn_status = request("cn_status")
etype = request("etype")
        
i = 1

if request("job_from") <> "" then
   job_from = request("job_from")
else
   job_from = chkdate(date())
end if

if request("job_to") <> "" then
   job_to = request("job_to")
else
   job_to = chkdate(date())
end if

if etype="" then
    etype="All"
End if 

sql2= "SELECT b.jobp_partcode,c.md_desc,sum(jobp_qty) as qty,(select TOP 1 md_averagecost from tblmodel_avgcost where md_code = b.jobp_partcode and CAST(tblmodel_avgcost.md_date as date) <= '" & ChkDateYYYYMMDD(job_to) & "' order by md_date desc) as avgcost" & _
	    " FROM tbljob_parts b " & _ 
	" inner join tbljob a on b.job_code = a.job_code "  & _ 
	" inner join tblmodel c on b.jobp_partcode = c.md_code " & _ 
		" where a.job_id is not null and a.job_status='Posted' " & _ 
    	" and cast(a.job_posteddate as date) >= '" & ChkDateYYYYMMDD(job_from) & "' and cast(a.job_posteddate as date) <= '" & ChkDateYYYYMMDD(job_to) & "' and job_actual_wrty_status='Under' "
		
if etype <> "All" then 
		sql2 = sql2 & " and c.md_type='" & etype & "' "
end if
 
sql2 = sql2 & " group by b.jobp_partcode,c.md_desc"


response.Cookies("GAPS")("sqlexcel") = sql2
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
if rs.eof then
   norecord = "There is no record found."
end if

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

link = "&orderby=" & orderby & "&ordertype=" & ordertype & "&etype=" & etype & "&cn_no=" & cn_no & "&job_from=" & job_from & "&job_to=" & job_to 
%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Under Warranty Spare Parts Summary </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1">&nbsp;</td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_under_wrty_parts_excel.asp?job_to=<%=job_to%>&etype=<%=etype%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_under_wrty_parts.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong>Job Posted Date<br />
                            </strong></td>
                        <td colspan="3"><div align="left"><strong><font color="#000000"><strong>
                          <input name="job_from" type="text" id="job_from" value="<%=job_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_to" type="text" id="job_to" value="<%=job_to%>"
                                            size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </div></td>
                      </tr>
                      <tr>
                        <td class="titlegrey1">Type</td>
                        <td><span class="titlegrey1">
                          <select name="etype" id="etype">
                            <option value="All">All</option>
                            <option value="CF" <%if etype="CF" then response.write " selected"%>>Ceiling Fan</option>
                            <option value="WH" <%if etype="WH" then response.write " selected"%>>Water Heater</option>
                            
                          </select>
                        </span></td>
                        <td width="24%" align="center">&nbsp;</td>
                        <td width="23%" rowspan="2"><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1">&nbsp;</td>
                        <td>
                          <span class="titlegrey1">
                          &nbsp;</span></td>
                        <td width="24%" align="center" valign="top"><label for="inv_no"></label></td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font> of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_under_wrty_parts.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_under_wrty_parts.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF">No</font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF">Spare Part Name</font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF">Spare Part Code</font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF">Avg Cost</font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF">Qty</font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF">Total</font></td>
                    </tr>
                    
<%

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
    
    if not isnull(rs("avgcost")) then
        totalavgcost=rs("qty") * avgcost
    end if

%>
                   <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center"><%=j%></td>
                      <td align="left" nowrap="nowrap"><%=rs("md_desc")%></td>
                      <td align="left" nowrap="nowrap"><strong> <%=rs("jobp_partcode")%></strong></td>
                      <td align="left" nowrap="nowrap"><%=rs("avgcost")%></td>
                      <td align="left" nowrap="nowrap"><%=rs("qty")%></td>
                      <td align="left" nowrap="nowrap"><%=totalavgcost%></td>
                    </tr>
<%
totalqty = totalqty + rs("qty")
totalcost = totalcost + totalavgcost 
count = count + 1 
i = i + 1
rs.MoveNext
totalavgcost=0
Next
rs.Close

%>                 
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="4" align="right" bgcolor="#CCCCCC"></td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                      <td align="right" bgcolor="#CCCCCC"></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Grand Total</strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#999999"><strong><%=(totalqty)%></strong></td>
                      <td align="left" bgcolor="#999999"><strong><%=chknumber2(totalcost)%></strong></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font> of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_under_wrty_parts.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_under_wrty_parts.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
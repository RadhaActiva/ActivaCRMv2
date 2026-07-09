<!-- #include file="header.asp" -->
<%
job_tech_code = request.cookies("GAPS")("sloginid")

jobmonth = request("jobmonth")
jobyear = request("jobyear")

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

sql2 = "select id, tech_code,clink from tbltech_claim_forms where tech_code = '" & job_tech_code & "' order by id"
    
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
link = "&jobyear=" & jobyear & "&jobmonth=" & jobmonth & "&job_tech_code=" & job_tech_code
%> 


<script>
function DisplayReport() 
{
	document.form1.action = "tech_view_claim_form.asp";
	document.form1.submit();
}
</script> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                         <td class="titleblue1"><div align="left">View Approved Claim Form</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%"> <strong></strong></td>
                      <td width="20%" align="center" class="titlegrey1">
                      </td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="tech_view_claim_form.asp">
                  <tr>                    
                  <td width="50%" height="30" align="left" bgcolor="#FFFFFF">Technician : <%=job_tech_code%></td>
                  <td width="20%" height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                  <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>                
                  <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>              
                 </form>                
              <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="6" align="right" class="style1"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='tech_view_claim_form.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='tech_view_claim_form.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="4%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="8%" align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Claim Form</span></strong></font></td>                      
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
     
%>                   
                   <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center"><%=j%></td>                             
                      <td align="left"><font color="#0000FF"><a href="<%=rs("clink")%>" target="_blank"><%=rs("clink")%></a></font></td>                       
                    </tr>
<%
   
count = count + 1 
i = i + 1
rs.MoveNext
Next

rs.Close
Set rs = Nothing

'different technicians type has diff logic for footer figures

%>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="6" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='tech_view_claim_form.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='tech_view_claim_form.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                </table></td>
                </tr>
                <tr>
                  <td height="30" colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
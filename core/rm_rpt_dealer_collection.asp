<!-- #include file="header.asp" -->
<head>
</head>
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
cust_code = request("cust_code")

if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-30,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

sql2= "select job_code,job_cust_code,job_cust_name, job_date, job_model, job_faulty_reason_cs,job_model_desc,job_tech_SN, job_type,job_tech_code ,job_donedate,job_posteddate,job_totallabourAmt,job_tech_service_date,job_actual_wrty_status,job_status,job_totalAmt,job_submitteddate,job_SN_no,job_dealer_inv from tbljob a where job_faulty_reason_cs like '%Installation%' " & _
	   "and job_posteddate >=  '" & ChkDateYYYYMMDD(job_date_from) & "' and job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' and a.job_dealer = '" & cust_code & "' and a.job_status in ('Posted') and a.job_payee='Dealer'"

'if request("cust_code") <> "" then 
'  sql2 = sql2 & " and a.job_dealer = '" & cust_code & "'"
'else
'    sql2 = sql2
'end if
'response.write sql2
    
response.Cookies("GAPS")("sqlexcel") = sql2
collection_amt = 0
labouramt = 0
t_totalpymt = 0
t_totalcoll = 0

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
link = "&orderby=" & orderby & "&dealer_code=" & cust_code & "&job_date_from=" & job_date_from  & "&job_date_to=" & job_date_to
%> 


<script>
function DisplayReport() 
{
	document.form1.action = "rm_rpt_dealer_collection.asp";
	document.form1.submit();
}
</script> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
              <table width="100%"><!--for UI Purpose-->
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Collection </font>Report By Dealer (Installation)</div>
                          </td>
                      </tr>
                    </table>
                    </td>
                </tr>
           
                <tr>
                    <br />
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                     
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="action_report.asp?type=warehouselocation">
                     <tr>
                  <td height="30" align="left" bgcolor="#FFFFFF" width="60%"><font color="#000000">Date Start &nbsp;<input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                  <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
                    calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong>
                 &nbsp; Date End &nbsp;<input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>" size="15" />
                  <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
                    calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong>&nbsp;&nbsp;&nbsp; 
                   <a href="rm_rpt_dealer_collection_excel.asp?dealer_code=<%=cust_code%>&job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                </tr>
                  <tr>
                   <td height="30" align="left" bgcolor="#FFFFFF" width="60%">Dealer&nbsp;                 
                    <select name="cust_code" id="cust_code">
                      <%			

                sql1= "SELECT DISTINCT a.cust_code, LTRIM(RTRIM(a.cust_name))    AS cust_name, LTRIM(RTRIM(a.cust_address)) AS cust_address " & _
                      "FROM tblcustomer a WHERE a.cust_type = 'Dealer' AND EXISTS " & _
                      "(SELECT 1 FROM tbljob b  WHERE b.job_dealer = a.cust_code and b.job_status in ('Posted')  AND b.job_faulty_reason_cs = 'Installation')"  & _
                       "ORDER BY LTRIM(RTRIM(a.cust_name)), LTRIM(RTRIM(a.cust_address));"
                set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  if (cust_code) = (rs1("cust_code")) then
					  response.write "<option value='" & rs1("cust_code") & "' selected>" & rs1("cust_code") & " - " & rs1("cust_name") & " - " & rs1("cust_address") & "</option>"
                      'wh_name = r1(wh_name)
                      else
					  response.write "<option value='" & rs1("cust_code") & "'>" & rs1("cust_code") & " - " & rs1("cust_name") & " - " & rs1("cust_address") & "</option>"
					  end if 					  
                     
				rs1.movenext
				wend
				rs1.close					
				%>
                    </select></td>
            <tr>
                  <td width="80%" height="30" align="left" bgcolor="#FFFFFF">&nbsp;&nbsp;
                    <span class="titlegrey1">
                    <input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" /> (only posted jobs & payee is dealer)</span></td>
            </tr>
                 </form>
                
              <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="7" align="right" class="style1"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_dealer_collection.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_dealer_collection.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                        <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Job Sheet No</span></strong></font></td>
                        <!--<td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Service Date</span></strong></font></td>-->
                        <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Dealer DO/Invoice No</span></strong></font></td>
                        <td width="343" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><span><strong>Model Desc</strong></span></font></td>
                        <td width="343" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><span><strong>Serial No</strong></span></font></td>
                      <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Warranty <br />Info</span></strong></font></td>
                        <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Type</span></strong></font></td>
                        <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Collection (MYR)</span></strong></font></td>
                      <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total <br /> Payment</span></strong></font></td>
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
                        <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("job_code")%></font></strong></td>
                       <td align="left" nowrap="nowrap"> <%=rs("job_dealer_inv")%></td>
                      <td align="left" nowrap="nowrap"> <%=rs("job_model_desc")%></td>
                       <td align="left" nowrap="nowrap"> <%=rs("job_SN_no")%></td>
                      <td align="left" nowrap="nowrap"><%=rs("job_actual_wrty_status")%></td>
                      <td align="left" nowrap="nowrap"><%=rs("job_type")%></td>
                        <%
                        labouramt = ChkNumber(rs("job_totallabourAmt"))

                        If Not IsNull(rs("job_type")) Then
                            If rs("job_type") = "WH" Then
                                collection_amt = labouramt - 20
                            ElseIf rs("job_type") = "CF" Then
                                collection_amt =labouramt - 10
                            End If
                            if collection_amt < 0 then
                                collection_amt = 0
                            End If
                        End If
                        %>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"> <%=chknumber2(collection_amt)%></td>
                        <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"> <%=chknumber2(rs("job_totalAmt"))%></td>
                      
                    </tr>
<%

t_totalcoll =  collection_amt + t_totalcoll
t_totalpymt = rs("job_totalAmt") + t_totalpymt 

count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
                   
                    <tr bgcolor="#F3F3F3">
                      <td height="30" colspan="7" align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(t_totalcoll)%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(t_totalpymt)%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="30" colspan="7" align="right" bgcolor="#999999"><strong>Total Job</strong> <strong>Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=j-1%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"></td>
                     </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="7" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_dealer_collection.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_dealer_collection.asp?num=" & Showed+row & link & "'> Next >></a>"
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
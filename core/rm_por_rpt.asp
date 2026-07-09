<!-- #include file="header.asp" -->
<%

if ordertype = "" then 
   ordertype = "desc"
end if

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
    

sql2= "select a.por_id,a.por_docno, a.por_date,por_part_code, por_eta1 as 'eta1_date', por_order_qty1,por_eta2 as 'eta2_date', por_order_qty2, [por_createdby],por_createddate as 'por_createddate' from tblpor A " & _
"where a.por_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and a.por_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' order by a.por_docno desc" 
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql2
rs1.CursorLocation  = 3
rs1.Open
if rs1.eof then
   norecord = "There is no record found."
end if

i = 1	
response.Cookies("GAPS")("sqlexcel") = sql2

If Not rs1.EOF Then

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs1.RecordCount
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

link = "&por_docno=" & por_docno  

%>  

<script>
function DisplayReport() 
{
	document.form1.action = "rm_por_rpt.asp";
	document.form1.submit();
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Parts </font>Order Plan (POP) Report</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1"></td>
                      <!--<td width="20%" align="center" class="titlegrey1"><a href="rm_por_rpt_excel.asp?por_docno=<%=por_docno%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>-->
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_por_rpt.asp">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td class="titlegrey1">POP Doc Date</td>
                        <td colspan="3"><strong><font color="#000000"><strong>
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
                        calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong>  eg: 21-May-2015 &nbsp;&nbsp;&nbsp<span class="titlegrey1">
                         <input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" />
                  </span></td></td>
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
					Response.Write " <a href='rm_por_rpt.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_por_rpt.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>POP Doc No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Doc Date</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Part #</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>ETA/ETS</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Order Qty</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>ETA/ETS</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Order Qty</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Created By</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Created Date</strong></font></td>
                    </tr>
                    
<%

if not rs1.eof then
rs1.Move Showed
count = Showed + 1
end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

     partdesc = ""
     sql = "select md_desc from tblmodel where md_code = '" & rs1("por_part_code") & "'  "
     partdesc = selectid(sql) 
%>                    
                    <tr bgcolor="<%=nbgcolor%>">
                      <td height="25" align="center" valign="top" nowrap="nowrap"><%=j%></td>
                      <td align="left" valign="top" nowrap="nowrap"><strong> <font color="#0000FF"><a href="rm_por_edit.asp?por_docno=<%=rs1("por_docno")%>&por_id=<%=rs1("por_id")%>&porreport=Yes" target="_blank"><%=rs1("por_docno")%></font></strong></td>                      
                      <td align="left" valign="top" nowrap="nowrap"><%=chkdate(rs1("por_date"))%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_part_code")%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("eta1_date")%></td>
				      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_order_qty1")%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("eta2_date")%></td>
				      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_order_qty2")%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_createdby")%></td>
                        <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_createddate")%></td>
                    </tr>
<%

count = count + 1 
i = i + 1
rs1.MoveNext
Next
rs1.Close

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
					Response.Write " <a href='rm_por_rpt.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_por_rpt.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
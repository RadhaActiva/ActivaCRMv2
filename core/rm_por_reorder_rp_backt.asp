<!-- #include file="header.asp" -->
<%

'sql2= "select a.por_id,a.por_docno, a.por_date,por_part_code, por_eta1 as 'eta1_date', por_order_qty1,por_eta2 as 'eta2_date', por_order_qty2, [por_createdby],por_createddate as 'por_createddate' from tblpor A " & _
'"where a.por_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and a.por_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' order by a.por_docno desc" 

Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = strconnect
cmd.CommandText = "sp_GetStockMonthCoverage"
cmd.CommandType = 4  ' adCmdStoredProc
cmd.Parameters.Append cmd.CreateParameter("@wh", 200, 1, 10, "W1")  ' warehouse code
Set rs1 = cmd.Execute

if rs1.eof then
   norecord = "There is no record found."
end if
       
i = 1	
'response.Cookies("GAPS")("sqlexcel") = sql2

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
	document.form1.action = "rm_por_reorder_rpt.asp";
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Parts </font>Order Plan (POP) Re-order Levels </div></td>
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
                  <td align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_por_reorder_rpt.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_por_reorder_rpt.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Spare Part</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Description</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Months</strong></font></td>   
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
   
%>                   
                 
                      <tr bgcolor="<%=nbgcolor%>">
                      <td height="25" align="center" valign="top" nowrap="nowrap"><%=j%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("ItemCode")%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("ItemDescription")%></td>
				      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("MonthsCover")%></td>
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
					Response.Write " <a href='rm_por_reorder_rpt.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_por_reorder_rpt.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
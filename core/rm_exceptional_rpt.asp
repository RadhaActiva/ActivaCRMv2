<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            color: #FFFFFF;
        }
    </style>
</head>
<%

etype = request("etype")
foundrecs=false      

i = 1



if etype = "DJOB" then 
      SQL2 = "select stk_itm_code, stk_voucherno,stk_logdate, count(*) as 'duplicate', 'Posted' as posted from tblstocktran where stk_type='Job' " & _
     "group by stk_itm_code, stk_voucherno, stk_logdate "  & _
     "HAVING COUNT(stk_voucherno) > 1 -- stk_itm_code stk_logdate"
      set rs = server.CreateObject("adodb.recordset")
      rs.ActiveConnection = strconnect
      rs.Source = sql2
      rs.CursorLocation  = 3
      rs.Open
      if rs.eof then
         norecord = "There is no record found."
      else
            foundrecs=true
      end if

If foundrecs=true  Then

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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Exceptional </font>Error Report Summary </div></td>
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
                      <td width="20%" align="center" class="titlegrey1"></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_exceptional_rpt.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">               
                      <tr>
                        <td class="titlegrey1">Exception Type</td>
                        <td><span class="titlegrey1">
                          <select name="etype" id="etype">                          
                            <option value="DJOB" <%if etype="DJOB" then response.write " selected"%>>Duplicate Job Sheet</option>                            
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
					Response.Write " <a href='rm_exceptional_rpt.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_exceptional_rpt.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                        <%If etype="DJOB" then %>
                              <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF">No</font></td>
                              <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF">Job No</font></td>
                              <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF">Stock Code</font></td>
                              <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF">Log Date</font></td>
                              <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"># of Duplicate</font></td>    
                              <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF">Status</font></td>  
                        <%end if %>
                    </tr>
                    
<%
If foundrecs=true  Then
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
                  <%If etype="DJOB" then %>
                      <td height="30" align="center"><%=i%></td>
                      <td align="left" nowrap="nowrap"><%=rs("stk_voucherno")%></td>
                      <td align="left" nowrap="nowrap"><strong> <%=rs("stk_itm_code")%></strong></td>
                      <td align="left" nowrap="nowrap"><%=rs("stk_logdate")%></td>
                      <td align="left" nowrap="nowrap"><%=rs("duplicate")%></td>   
                      <td align="left" nowrap="nowrap"><%=rs("posted")%></td> 
                   <%end if %>
                    </tr>
<%
   
    count = count + 1 
    i = i + 1


    rs.MoveNext
    Next
    rs.Close
end if
%>                 
       
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
					Response.Write " <a href='rm_exceptional_rpt.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_exceptional_rpt.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>

<!-- #include file="footer2.asp" -->
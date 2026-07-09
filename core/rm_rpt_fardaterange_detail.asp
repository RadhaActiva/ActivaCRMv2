<!-- #include file="database/datastore.asp" -->

<%
job_date_from = request("job_date_from")
job_date_to = request("job_date_to")
state_id = request("state_id")
job_tech_type = request("job_tech_type")
faulth_code=request("faulth_code") 
stype=request("stype") 

if request("job_tech_model") <> "" then
   job_tech_model = replace(request("job_tech_model"), " ", "")
   arrjob_tech_model = split(job_tech_model,",")
   job_tech_model = replace(job_tech_model, ",", "','")
   
   listjob_tech_model = listjob_tech_model & job_tech_model
   
else
   listjob_tech_model = ""
   arrjob_tech_model = split("0,0",",")
   
end if
%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>

<%
i = 1
sql2 = "SELECT tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, " & _
		"tbljob.job_cust_postcode, tbljob.job_cust_state, tbljob.job_cust_state_id, tbljob.job_cust_city, tbljob.job_cust_city_id, tbljob.job_cust_email,  " & _
		"tbljob.job_cust_tel1, tbljob.job_cust_tel2, tbljob.job_createddate, tbljob.job_createdby, tbljob.job_JS_receiveddate, tbljob.job_JS_receivedby,  " & _
		"tbljob.job_status, tbljob.job_purchase_date, tbljob.job_onlineWrtyNo, tbljob.job_onlineWrtyStatus, tbljob.job_type, tbljob.job_SN_no,  " & _
		"tbljob.job_Model, tbljob.job_Model_desc, tbljob.job_faulty_desc, tbljob.job_reportedby, tbljob.job_appointment_date, tbljob.job_appointment_time,  " & _
		"tbljob.job_tech_code, tbljob.job_appointment_remark, tbljob.job_emailsentdate, tbljob.job_emailsent, tbljob.job_smssentdate,  " & _
		"tbljob.job_smssent, tbljob.job_tech_type, tbljob.job_tech_model, tbljob.job_tech_tax_invoice, tbljob.job_tech_SN, tbljob.job_tech_faulty_code, " & _
		"tbljob.job_tech_faulty_reason, tbljob.job_tech_faulty_action, tbljob.job_tech_status, tbljob.job_tech_product_collectdate,  " & _
		"tbljob.job_tech_returntoCustDate, tbljob.job_actual_wrty_status, tbljob.job_wrty_photo, tbljob.job_hq_remark,  " & _
		"tbljob.job_hq_category_code, tbljob.job_hq_received_date, tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt, tbljob.job_totaltransportAmt,  " & _
		"tbljob.job_totalAmt, tbljob.job_repair_date, tbljob.job_return_tech_date, tbljob.job_office_issueRemark, tbljob.job_office_supervisor,  " & _
		"tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_inv_no, tbljob.job_do_no, tbljob.job_posteddate, tbltechnician.tech_name, tbltechnician.tech_tel1 " & _
		"FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null " & _
		"and tbljob.job_status='Posted' " & _
		"and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' "
		
	if faulth_code <> "" and faulth_code <> "all" then 	  
	   sql2 = sql2 & " and tbljob.job_tech_faulty_code = '" & faulth_code & "' "
	end if
	
	if state_id <> "" and state_id <> "0" then 	  
	   sql2 = sql2 & " and tbljob.job_cust_state_id = " & state_id & " "
	end if
		  
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
	end if
	
	if stype = "Over" then 
	   sql2 = sql2 & " and job_actual_wrty_status='Over' "
	elseif stype = "Under" then    
	   sql2 = sql2 & " and job_actual_wrty_status='Under' "
	elseif stype = "fa_MD_over" then  
	   sql2 = sql2 & " and job_actual_wrty_status='Over' and job_hq_category_code='MD' "    
	elseif stype = "fa_MD_under" then  
	   sql2 = sql2 & " and job_actual_wrty_status='Under' and job_hq_category_code='MD' "      
	elseif stype = "fa_DS_over" then  
	   sql2 = sql2 & " and job_actual_wrty_status='Over' and job_hq_category_code='DS' "   
	elseif stype = "fa_DS_under" then  
	   sql2 = sql2 & " and job_actual_wrty_status='Under' and job_hq_category_code='DS' "    
	elseif stype = "fa_WI_over" then  
	   sql2 = sql2 & " and job_actual_wrty_status='Over' and job_hq_category_code='WI' "       
	elseif stype = "fa_WI_under" then  
	   sql2 = sql2 & " and job_actual_wrty_status='Under' and job_hq_category_code='WI' "  
	elseif stype = "fa_CF_over" then  
	   sql2 = sql2 & " and job_actual_wrty_status='Over' and job_hq_category_code='CF' " 
	elseif stype = "fa_CF_under" then  
	   sql2 = sql2 & " and job_actual_wrty_status='Under' and job_hq_category_code='CF' "                          
	end if

sql2 = sql2 & " order by tbljob.job_code"

'response.write sql2
set rs1 = server.CreateObject("adodb.recordset")
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

link = "&job_tech_type=" & job_tech_type & "&faulth_code=" & faulth_code & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to & "&job_tech_model=" & job_tech_model & "&stype=" & stype & "&state_id=" & state_id 
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td width="616" class="style21"><font size="4"><strong>Job Listing</strong></font></td>
    <td width="142" align="center" class="style21"><a href="rm_rpt_fardaterange_detail_excel.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&faulth_code=<%=faulth_code%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=<%=stype%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0"></a></td>
  </tr>
  <tr> 
    <td colspan="2" align="right" valign="top"><strong>Page</strong> <font color="3366ff"> 
      <%=pagestartno%>
      </font>of <font color="3366ff"> 
      <%=pgCount%>
      </font>: 
      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_fardaterange_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_fardaterange_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td colspan="2" valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td><strong>Job No</strong></td>
          <td><strong>Status</strong></td>
          <td><strong>Job Date</strong></td>
          <td><strong>Posted Date</strong></td>
          <td align="left" class='tktTotals'><strong>Model</strong></td>
          <td align="left" class='tktTotals'><strong>Serial</strong></td>
          <td class='tktTotals'><strong>Wrty</strong></td>
          <td class='tktTotals'><strong>Faulty</strong></td>
          <td class='tktTotals'><strong>Faulty Reason</strong></td>
          <td class='tktTotals'><strong>Category Code</strong></td>
          <td class='tktTotals'><strong>Spare Part</strong></td>
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
        
        <tr valign="top" bgcolor="<%=nbgcolor%>"> 
          <td nowrap><%=count%>.</td>
          <td nowrap><%=rs("job_code")%></td>
          <td><%=rs("job_status")%></td>
          <td><%=chkdate(rs("job_date"))%></td>
          <td nowrap><%=chkdate(rs("job_posteddate"))%></td>
          <td align="left" nowrap><%=rs("job_tech_model")%></td>
          <td align="left" nowrap><%=rs("job_tech_SN")%></td>
          <td><%=rs("job_actual_wrty_status")%></td>
          <td><%=rs("job_tech_faulty_code")%></td>
          <td><%=rs("job_tech_faulty_reason")%></td>
          <td><%=rs("job_hq_category_code")%></td>
          <td><%
		        k= 1
				sql1 = "SELECT jobp_id, job_code, jobp_partcode, jobp_desc, jobp_unitcost, jobp_discountamt, jobp_discounttype, " & _
				      "jobp_netcost, jobp_qty, jobp_subtotal FROM tbljob_parts where job_code='" & rs("job_code") & "'"	
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
			
			%>
            <table width="100%" border="1" cellspacing="0" cellpadding="3">
              <tr>
                <td><%=rs1("jobp_partcode")%></td>
                <td><%=rs1("jobp_desc")%></td>
                <td><%=rs1("jobp_qty")%></td>
                <td><%=rs1("jobp_netcost")%></td>
              </tr>
            </table>
          <%	
				k = k + 1
				rs1.movenext
				wend
				rs1.close
			  %></td>
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
  <tr valign="top"> 
    <td colspan="10" align="right"><strong>Page</strong> <font color="3366ff"> 
      <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_fardaterange_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_fardaterange_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

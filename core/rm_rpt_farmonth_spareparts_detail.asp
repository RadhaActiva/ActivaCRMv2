<!-- #include file="database/datastore.asp" -->

<%
job_tech_type = request("job_tech_type")
parts_code=request("parts_code") 
jobmonth=request("jobmonth") 
jobyear=request("jobyear") 
stype=request("stype") 
lastmonth=request("lastmonth") 

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
		"tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_inv_no, tbljob.job_do_no " & _
		"FROM tbljob_parts INNER JOIN tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	    "where tbljob.job_id is not null " & _ 
		"and tbljob_parts.jobp_partcode='" & parts_code & "' and tbljob.job_status='Posted' " 
		
	if jobmonth <> "0" then 	  
	   sql2 = sql2 & " and month(tbljob.job_posteddate) = '" & jobmonth & "' and  year(tbljob.job_posteddate) = '" & jobyear & "' " 
	end if
	
	if lastmonth<>"" then
	   currentdate="1-" & convertmonth(lastmonth) & "-" & jobyear 
	   
	   fromdatemonth=dateadd("m",-11, currentdate)
	   todatemonth=dateadd("m",1, currentdate)
	   
	   sql2 = sql2 & " and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(chkdate(fromdatemonth)) & "' and tbljob.job_posteddate < '" & ChkDateYYYYMMDD(chkdate(todatemonth)) & "' "
	end if
		  
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model & "') "
	end if
	
	if stype = "Over" then 
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Over' "
	elseif stype = "Under" then    
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Under' "
	elseif stype = "fa_MD_over" then  
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Over' and job_hq_category_code='MD' "    
	elseif stype = "fa_MD_under" then  
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Under' and job_hq_category_code='MD' "      
	elseif stype = "fa_DS_over" then  
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Over' and job_hq_category_code='DS' "   
	elseif stype = "fa_DS_under" then  
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Under' and job_hq_category_code='DS' "    
	elseif stype = "fa_WI_over" then  
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Over' and job_hq_category_code='WI' "       
	elseif stype = "fa_WI_under" then  
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Under' and job_hq_category_code='WI' "  
	elseif stype = "fa_CF_over" then  
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Over' and job_hq_category_code='CF' " 
	elseif stype = "fa_CF_under" then  
	   sql2 = sql2 & " and tbljob.job_actual_wrty_status='Under' and job_hq_category_code='CF' "                          
	end if

sql2 = sql2 & " order by tbljob.job_code"

''response.write sql2
'response.End()

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

link = "&job_tech_type=" & job_tech_type & "&parts_code=" & parts_code & "&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&job_tech_model=" & request("job_tech_model") & "&stype=" & stype & "&lastmonth=" & lastmonth 
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Job Listing</strong></font></td>
  </tr>
  <tr> 
    <td align="right" valign="top"><strong>Page</strong> <font color="3366ff"> 
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
					Response.Write " <a href='rm_rpt_farmonth_spareparts_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_farmonth_spareparts_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td><strong>Job No</strong></td>
          <td><strong>Status</strong></td>
          <td><strong>Job Date</strong></td>
          <td align="left" class='tktTotals'><strong>Model</strong></td>
          <td class='tktTotals'><strong>Wrty</strong></td>
          <td class='tktTotals'><strong>Faulty</strong></td>
          <td class='tktTotals'><strong>Part Code</strong></td>
          <td class='tktTotals'><strong>Category Code</strong></td>
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
          <td nowrap><%=chkdate(rs("job_date"))%></td>
          <td align="left" nowrap><%=rs("job_tech_model")%></td>
          <td><%=rs("job_actual_wrty_status")%></td>
          <td><%=rs("job_tech_faulty_code")%></td>
          <td><%=rs("job_tech_faulty_reason")%></td>
          <td><%=rs("job_hq_category_code")%></td>
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
    <td colspan="9" align="right"><strong>Page</strong> <font color="3366ff"> 
      <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_farmonth_spareparts_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_farmonth_spareparts_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

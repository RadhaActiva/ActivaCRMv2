<!-- #include file="database/datastore.asp" -->

<%
set rs = server.CreateObject("adodb.recordset")

if request("job_code") <> "" then	  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo, job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
		"job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, job_repair_date, job_return_tech_date,  " & _
		"job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_inv_date, job_do_no, job_do_date, job_submittedby, job_submitteddate, " & _
		"job_doneby, job_donedate, job_postedby, job_posteddate, job_cancelledby, job_cancelleddate  " & _
	    "FROM tbljob WHERE job_code = '" & request("job_code") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			job_id = rs("job_id") 
			job_code = rs("job_code") 
			job_date = rs("job_date")
			job_cust_code = rs("job_cust_code")
			job_cust_name = rs("job_cust_name")
			job_cust_address = rs("job_cust_address") 
			job_cust_postcode = rs("job_cust_postcode") 
			job_cust_state = rs("job_cust_state") 
			job_cust_state_id = rs("job_cust_state_id") 
			job_cust_city = rs("job_cust_city") 
			job_cust_email = rs("job_cust_email") 
			job_cust_tel1 = rs("job_cust_tel1") 
			job_cust_tel2 = rs("job_cust_tel2")  
			job_remark = rs("job_remark")
			job_createddate = rs("job_createddate") 
			job_createdby = rs("job_createdby") 
			job_JS_receiveddate = rs("job_JS_receiveddate") 
			job_JS_receivedby = rs("job_JS_receivedby") 
			job_status = rs("job_status") 
			job_purchase_date = rs("job_purchase_date") 
			job_onlineWrtyNo = rs("job_onlineWrtyNo") 
			job_onlineWrtyStatus = rs("job_onlineWrtyStatus") 
			job_type = rs("job_type") 
			job_SN_no = rs("job_SN_no") 
			job_Model = rs("job_Model") 
			job_model_desc = rs("job_model_desc") 
			job_faulty_reason_cs = rs("job_faulty_reason_cs") 
			job_faulty_desc = rs("job_faulty_desc") 
			job_reportedby = rs("job_reportedby") 
			job_appointment_date = rs("job_appointment_date") 
			job_appointment_time = rs("job_appointment_time") 
			job_tech_code = rs("job_tech_code") 
			job_appointment_remark = rs("job_appointment_remark") 
			job_emailsentdate = rs("job_emailsentdate") 
			job_emailsent = rs("job_emailsent") 
			job_smssentdate = rs("job_smssentdate") 
			job_smssent = rs("job_smssent") 
			job_tech_type = rs("job_tech_type") 
			job_tech_model = rs("job_tech_model") 
			job_tech_tax_invoice = rs("job_tech_tax_invoice") 
			job_tech_SN = rs("job_tech_SN") 
			job_tech_faulty_reason = rs("job_tech_faulty_reason") 
			job_tech_faulty_action = rs("job_tech_faulty_action") 
			job_tech_status = rs("job_tech_status") 
			job_tech_product_collectdate = rs("job_tech_product_collectdate") 
			job_tech_returntoCustDate = rs("job_tech_returntoCustDate") 
			job_actual_wrty_status = rs("job_actual_wrty_status") 
			job_wrty_photo = rs("job_wrty_photo") 
			job_tech_logby  = rs("job_tech_logby")
			job_tech_logdate  = rs("job_tech_logdate")
			job_hq_remark = rs("job_hq_remark") 
			job_hq_category_code = rs("job_hq_category_code") 
			job_hq_received_date = rs("job_hq_received_date") 
			job_totalPartsAmt = rs("job_totalPartsAmt") 
			job_totallabourAmt = rs("job_totallabourAmt") 
			job_totaltransportAmt = rs("job_totaltransportAmt") 
			job_totalAmt = rs("job_totalAmt") 
			job_repair_date = rs("job_repair_date") 
			job_return_tech_date = rs("job_return_tech_date") 
			job_office_issueRemark = rs("job_office_issueRemark") 
			job_office_supervisor = rs("job_office_supervisor") 
			job_office_taxinvoice = rs("job_office_taxinvoice") 
			job_rcn_no = rs("job_rcn_no")
			job_rcn_Date = rs("job_rcn_Date")
			job_inv_no = rs("job_inv_no") 
			job_inv_date = rs("job_inv_date") 
			job_do_no = rs("job_do_no")
			job_do_date = rs("job_do_date") 
			job_submittedby = rs("job_submittedby")
			job_submitteddate = rs("job_submitteddate")
			job_doneby = rs("job_doneby")
			job_donedate = rs("job_donedate")
			job_postedby = rs("job_postedby")
			job_posteddate = rs("job_posteddate")
			job_cancelledby = rs("job_cancelledby")
			job_cancelleddate = rs("job_cancelleddate")
		End If
        rs.close


		if job_tech_code <> "" then	  
		sql = "SELECT tech_id, tech_code, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
			  "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id " & _
			  "FROM tbltechnician WHERE tech_code = '" & job_tech_code & "' "
				rs.Open sql,strconnect,0,1,&H0001
				If Not rs.EOF Then
					job_emailsent = rs("tech_email") 
					job_smssent = rs("tech_tel1") 
				End If
				rs.Close
		end if
		
		if job_cust_state_id <> "" then 
		sql = "select state_code from tblstate where state_id=" & job_cust_state_id
		state_code = selectid(sql)
		end if

end if
 
	
Dim dayPart, monthPart, yearPart,formattedDate
dayPart = Day(job_date)
monthPart = MonthName(Month(job_date), True)
yearPart = Right(Year(job_date), 2)
formattedDate = Right("0" & dayPart, 2) & monthPart & yearPart

job_model_desc = replace(job_model_desc, "RUBINE","")
job_model_desc = replace(job_model_desc, "INSTANT WATER HEATER","")
job_model_desc = replace(job_model_desc, "CEILING FAN","")
job_model_desc = replace(job_model_desc, "`","")
job_cust_tel1 = replace(job_cust_tel1, "-","")
job_cust_tel1 = "0" + job_cust_tel1
	  
'smscontent = "RIEGEN " & "J:" & job_code & chr(13) & " " & formattedDate & chr(13) & job_cust_name & chr(13) & "T:" & job_cust_tel1 & "/" & job_cust_tel2 & chr(13) & _
'             "M:" & job_model_desc & chr(13) & job_cust_address & " " & job_cust_city & " " & _
'			  job_cust_postcode & " " & state_code

smscontent = "RIEGEN " & "J:" & job_code & chr(13) & " " & formattedDate & chr(13) & job_cust_name & chr(13) & "T:" & job_cust_tel1 & "/" & job_cust_tel2 & chr(13) & _
            "M:" & job_model_desc & chr(13) & job_cust_address & " " & job_cust_city & " " & _
			  job_cust_postcode & " " & state_code
smsnumber = request("smsnumber")
%>

    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
    <html>
    <head>
    </head>
    
    <body onLoad="document.forms['form1'].submit();">
	<form name="form1" method="post" action="https://broadcast.smsgateway.cc/API/ReceiveSMS.aspx?Username=irene.tey%40riegen.com.my&Password=n9bX%23hUE%5E%21ed%21We&MobileNumber=<%=smsnumber%>&ContentType=1&MsgText=<%=smscontent%>">
<br>
<br>
<br>
<br>

    <input type="submit" name="Submit" value="Submit">
    </form>
    </body>
    </html>

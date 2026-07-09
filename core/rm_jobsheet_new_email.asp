<!-- #include file="database/datastore.asp" -->

<%
set rs = server.CreateObject("adodb.recordset")

if request("job_code") <> "" then	  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city, job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_model_desc, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_code, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_service_date, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo, job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
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
			job_tech_model_desc = rs("job_tech_model_desc") 
			job_tech_tax_invoice = rs("job_tech_tax_invoice") 
			job_tech_SN = rs("job_tech_SN") 
			job_tech_faulty_code = rs("job_tech_faulty_code") 
			job_tech_faulty_reason = rs("job_tech_faulty_reason") 
			job_tech_faulty_action = rs("job_tech_faulty_action") 
			job_tech_status = rs("job_tech_status") 
			job_tech_product_collectdate = rs("job_tech_product_collectdate") 
			job_tech_service_date = rs("job_tech_service_date") 
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
end if
 
Subject = CompanyName & " Riegen Marketing System - Notification for Job: " & job_code & ", Customer Code: " & job_cust_code & " - " & job_cust_name  
sbody = "" & _
 "<html> " & _
 "<head> " & _
 "<title>" & CompanyName & "- Riegen Marketing System</title> " & _
 "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'> " & _
 "</head> " & _
 "<body> <br>" & _
 "  <p>Dear<strong> " & job_cust_name & " </strong>,<br /> " & _
 "  <br /> " & _
 "  This is a  notice that an Job Sheet has been generated on " & chkdate(job_date) & "<br /> " & _
 "  <br /> " & _
 "  <strong>Job Code:   " & job_code & "</strong><br /> " & _
 "  <strong>Job Remark:   " & request("email_remark") & "</strong><br /> " & _ 
 "  <br /> " & _
 "  Thank you for your support.<br /> " & _
 "  <p><br /> " & _
 "  Regards,<br /> " & _
 "  <strong>Customer Service.</strong><br /> " & _
 "  <br /> " & _
 "  <strong>" & CompanyHeaderName & " " & CompanyHeaderReg & "</strong><br /> " & _
 "  " & CompanyHeaderAddress & " <br /> " & _
 "    Hunting Line TEL: " & CompanyHeaderTel & "<br /> " & _
 "  " & CompanyHeaderWeb & " " & _
 "  </p> " & _
 "</p> " & _
 "</body> " & _
 "</html> "


'''Generate PDF
Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument
Doc.ImportFromUrl ImportFromUrlpath & "rm_jobsheet_new_print_pdf.asp?job_code=" & job_code
Filename = Doc.Save(Server.MapPath(documentpath & job_code & ".pdf"), true )
documents = Server.MapPath(documentpath & job_code & ".pdf")

	
emailto = request("emailto") & ","
listemail = split(emailto, ",")


for i = 0 to ubound(listemail)

	if instr(listemail(i),"@") > 0 then
		sendemailA "admin@riegen.com.my" ,listemail(i) ,Subject ,sbody, documents	
		response.write "Email Subject: <b>[" & Subject & "]</b> has sent to " & listemail(i) & "<br>"
    end if

next

'sendemailA Response.Cookies("GAPS")("email"), Response.Cookies("GAPS")("email") ,Subject ,sbody, documents


response.write "Process end."
%>
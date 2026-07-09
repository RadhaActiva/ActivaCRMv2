<!-- #include file="database/datastore.asp" -->

<%
set rs = server.CreateObject("adodb.recordset")

if request("do_no") <> "" then	  
sql = "SELECT do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, do_cust_state, " & _
		"do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_remark, do_createddate, do_createdby, do_job_code,  " & _
		"do_tech_code, do_totalqty, do_labourAmt, do_transportAmt, do_totalAmt, do_emailsent, do_emailsentdate, do_deliveredby, do_delivereddate,  " & _
		"do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate, " & _
		"do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, " & _
		"do_appointment_date, do_appointment_time, do_appointment_remark " & _
		"FROM tbldo WHERE do_no = '" & request("do_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			do_id = rs("do_id") 
			do_no = rs("do_no") 
			do_status = rs("do_status")
			do_date = rs("do_date")
			do_inv_no = rs("do_inv_no") 
			do_inv_date = rs("do_inv_date") 
			do_cust_code = rs("do_cust_code")
			do_cust_name = rs("do_cust_name")
			do_cust_address = rs("do_cust_address") 
			do_cust_postcode = rs("do_cust_postcode") 
			do_cust_state = rs("do_cust_state") 
			do_cust_state_id = rs("do_cust_state_id") 
			do_cust_city = rs("do_cust_city") 
			do_cust_city_id = rs("do_cust_city_id") 
			do_cust_email = rs("do_cust_email") 
			do_cust_tel1 = rs("do_cust_tel1") 
			do_cust_tel2 = rs("do_cust_tel2")  
			do_remark = rs("do_remark")
			do_createddate = rs("do_createddate") 
			do_createdby = rs("do_createdby") 
			do_tech_code = rs("do_tech_code") 
			do_totalqty = rs("do_totalqty") 
			do_labourAmt = rs("do_labourAmt")
			do_transportAmt = rs("do_transportAmt")
			do_totalAmt = rs("do_totalAmt")
			do_emailsent = rs("do_emailsent")
			do_emailsentdate = rs("do_emailsentdate")
			do_deliveredby = rs("do_deliveredby")
			do_delivereddate = rs("do_delivereddate")
			do_doneby = rs("do_doneby")
			do_donedate = rs("do_donedate")
			do_postedby = rs("do_postedby")
			do_posteddate = rs("do_posteddate")
			do_cancelledby = rs("do_cancelledby")
			do_cancelleddate = rs("do_cancelleddate")
			
			do_purchase_date = rs("do_purchase_date")
			do_onlineWrtyNo = rs("do_onlineWrtyNo")
			do_onlineWrtyStatus = rs("do_onlineWrtyStatus")
			do_SN_no = rs("do_SN_no")
			do_type = rs("do_type")
			do_Model = rs("do_Model")
			do_model_desc = rs("do_model_desc")
			
			do_appointment_date = rs("do_appointment_date")
			do_appointment_time = rs("do_appointment_time")
			do_appointment_remark = rs("do_appointment_remark")
			
			
		End If
		rs.Close
end if  
 
Subject = CompanyName & " Riegen Marketing System - Notification for DO: " & do_code & ", Customer Code: " & do_cust_code & " - " & do_cust_name  
sbody = "" & _
 "<html> " & _
 "<head> " & _
 "<title>" & CompanyName & "- Riegen Marketing System</title> " & _
 "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'> " & _
 "</head> " & _
 "<body> " & request("email_remark") & "<br>" & _
 "  <p>Dear<strong> " & do_cust_name & " </strong>,<br /> " & _
 "  <br /> " & _
 "  This is a  notice that an invoice has been generated on " & chkdate(do_date) & "<br /> " & _
 "  <br /> " & _
 "  <strong>Invoice  " & do_code & "</strong><br /> " & _
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
Doc.ImportFromUrl ImportFromUrlpath & "rm_do_print_pdf.asp?do_code=" & do_code
Filename = Doc.Save(Server.MapPath(documentpath & do_code & ".pdf"), true )
documents = Server.MapPath(documentpath & do_code & ".pdf")

'documents = Server.MapPath("/liveReports/doc/" & do_code & ".pdf")
	
emailto = request("emailto") & ","
listemail = split(emailto, ",")


for i = 0 to ubound(listemail)

	if instr(listemail(i),"@") > 0 then
		sendemailA Response.Cookies("GAPS")("email") ,listemail(i) ,Subject ,sbody, documents	
		response.write "Email Subject: <b>[" & Subject & "]</b> has sent to " & listemail(i) & "<br>"
    end if

next

sendemailA Response.Cookies("GAPS")("email"), Response.Cookies("GAPS")("email") ,Subject ,sbody, documents


response.write "Process end."
%>
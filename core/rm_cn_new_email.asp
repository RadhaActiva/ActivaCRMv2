<!-- #include file="database/datastore.asp" -->

<%
set rs = server.CreateObject("adodb.recordset")

if request("cn_no") <> "" then	  
sql = "SELECT cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
	  "cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
	  "cn_job_code, cn_do_no, cn_invoice_no, cn_totalqty, cn_totalPartsAmt, cn_remark, cn_labourAmt, cn_transportAmt, cn_gstAmt, cn_totalAmt,  " & _
	  "cn_emailsent, cn_emailsentdate, cn_returnedby, cn_returneddate, cn_submittedby, cn_submitteddate, cn_doneby, cn_donedate, cn_postedby,  " & _
	  "cn_posteddate, cn_cancelledby, cn_cancelleddate " & _
	  "FROM tblcn WHERE cn_no = '" & request("cn_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			cn_id = rs("cn_id") 
			cn_no = rs("cn_no") 
			cn_status = rs("cn_status")
			cn_date = rs("cn_date") 
			cn_inv_no = rs("cn_inv_no")  
			cn_inv_date = rs("cn_inv_date")  
			cn_cust_code = rs("cn_cust_code") 
			cn_cust_name = rs("cn_cust_name")  
			cn_cust_address = rs("cn_cust_address")  
			cn_cust_postcode = rs("cn_cust_postcode")  
			cn_cust_state = rs("cn_cust_state")  
			cn_cust_state_id = rs("cn_cust_state_id")  
			cn_cust_city = rs("cn_cust_city")  
			cn_cust_city_id = rs("cn_cust_city_id") 
			cn_cust_email = rs("cn_cust_email")  
			cn_cust_tel1 = rs("cn_cust_tel1")  
			cn_cust_tel2 = rs("cn_cust_tel2")  
			cn_createddate = rs("cn_createddate")  
			cn_createdby = rs("cn_createdby")  
			cn_job_code = rs("cn_job_code")  
			cn_do_no = rs("cn_do_no")  
			cn_invoice_no = rs("cn_invoice_no")  
			cn_totalqty = rs("cn_totalqty")  
			cn_totalPartsAmt = rs("cn_totalPartsAmt")  
			cn_remark = rs("cn_remark")  
			cn_labourAmt = rs("cn_labourAmt")  
			cn_transportAmt = rs("cn_transportAmt")  
			cn_gstAmt = rs("cn_gstAmt")  
			cn_totalAmt = rs("cn_totalAmt")  
			cn_emailsent = rs("cn_emailsent")  
			cn_emailsentdate = rs("cn_emailsentdate")  
			cn_returnedby = rs("cn_returnedby")  
			cn_returneddate = rs("cn_returneddate")  
			cn_submittedby = rs("cn_submittedby")  
			cn_submitteddate = rs("cn_submitteddate")  
			cn_doneby = rs("cn_doneby")  
			cn_donedate = rs("cn_donedate")  
			cn_postedby = rs("cn_postedby")  
			cn_posteddate = rs("cn_posteddate")  
			cn_cancelledby = rs("cn_cancelledby")  
			cn_cancelleddate = rs("cn_cancelleddate")  
		End If
end if
 
Subject = CompanyName & " Riegen Marketing System - Notification for CN: " & cn_no & ", Customer Code: " & cn_cust_code & " - " & cn_cust_name  
sbody = "" & _
 "<html> " & _
 "<head> " & _
 "<title>" & CompanyName & "- Riegen Marketing System</title> " & _
 "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'> " & _
 "</head> " & _
 "<body> <br>" & _
 "  <p>Dear<strong> " & cn_cust_name & " </strong>,<br /> " & _
 "  <br /> " & _
 "  This is a  notice that an CN has been generated on " & chkdate(cn_date) & "<br /> " & _
 "  <br /> " & _
 "  <strong>CN:   " & cn_no & "</strong><br /> " & _
 "  <strong>CN Remark:   " & request("email_remark") & "</strong><br /> " & _ 
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
Doc.ImportFromUrl ImportFromUrlpath & "rm_cn_new_print_pdf.asp?cn_no=" & cn_no
Filename = Doc.Save(Server.MapPath(documentpath & cn_no & ".pdf"), true )
documents = Server.MapPath(documentpath & cn_no & ".pdf")

	
emailto = request("emailto") & ","
listemail = split(emailto, ",")


for i = 0 to ubound(listemail)

	if instr(listemail(i),"@") > 0 then
		sendemailA "crm@riegen.com.my" ,listemail(i) ,Subject ,sbody, documents	
		response.write "Email Subject: <b>[" & Subject & "]</b> has sent to " & listemail(i) & "<br>"
    end if

next

'sendemailA Response.Cookies("GAPS")("email"), Response.Cookies("GAPS")("email") ,Subject ,sbody, documents


response.write "Process end."
%>
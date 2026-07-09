<!-- #include file="database/datastore.asp" -->

<%
set rs = server.CreateObject("adodb.recordset")

if request("inv_no") <> "" then	  
sql = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
		"inv_cust_city, inv_cust_city_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
		"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
		"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_remark, inv_job_code  " & _
		"FROM tblinvoice WHERE inv_no = '" & request("inv_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			inv_id = rs("inv_id") 
			inv_no = rs("inv_no") 
			inv_date = rs("inv_date") 
			inv_cust_code = rs("inv_cust_code")
			inv_cust_name = rs("inv_cust_name")
			inv_cust_address = rs("inv_cust_address") 
			inv_cust_postcode = rs("inv_cust_postcode") 
			inv_cust_state = rs("inv_cust_state") 
			inv_cust_state_id = rs("inv_cust_state_id") 
			inv_cust_city = rs("inv_cust_city") 
			inv_cust_city_id = rs("inv_cust_city_id") 
			inv_cust_email = rs("inv_cust_email") 
			inv_cust_tel1 = rs("inv_cust_tel1") 
			inv_cust_tel2 = rs("inv_cust_tel2")  
			inv_createddate = rs("inv_createddate") 
			inv_createdby = rs("inv_createdby") 
			inv_approveddate = rs("inv_approveddate") 
			inv_approvedby = rs("inv_approvedby") 
			inv_status = rs("inv_status") 
			inv_job_code = rs("inv_job_code")
			
			inv_gstAmt = rs("inv_gstAmt")
			inv_totalPartsAmt = rs("inv_totalPartsAmt")
			inv_totalAmt = rs("inv_totalAmt")
			inv_remark = rs("inv_remark")
		End If
		rs.Close
end if
 
Subject = CompanyName & " Riegen Marketing System - Notification for Invoice: " & inv_no & ", Customer Code: " & inv_cust_code & " - " & inv_cust_name  
sbody = "" & _
 "<html> " & _
 "<head> " & _
 "<title>" & CompanyName & "- Riegen Marketing System</title> " & _
 "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'> " & _
 "</head> " & _
 "<body> <br>" & _
 "  <p>Dear<strong> " & inv_cust_name & " </strong>,<br /> " & _
 "  <br /> " & _
 "  This is a  notice that an Invoice has been generated on " & chkdate(inv_date) & "<br /> " & _
 "  <br /> " & _
 "  <strong>Invoice No:   " & inv_no & "</strong><br /> " & _
 "  <strong>Invoice Remark:   " & request("email_remark") & "</strong><br /> " & _ 
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
Doc.ImportFromUrl ImportFromUrlpath & "rm_invoice_new_print_pdf.asp?inv_no=" & inv_no
Filename = Doc.Save(Server.MapPath(documentpath & inv_no & ".pdf"), true )
documents = Server.MapPath(documentpath & inv_no & ".pdf")

	
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
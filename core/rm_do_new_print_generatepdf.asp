
<!-- #include file="database/dbconnect.asp" -->

<%

'''Generate PDF

Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument
Doc.ImportFromUrl ImportFromUrlpath & "rm_do_new_print_pdf.asp?do_no=" & request("do_no")

Filename = Doc.Save(Server.MapPath(documentpath & request("do_no") & ".pdf"), true )	
response.redirect ImportFromUrlpath & "/shared/" & request("do_no") & ".pdf"
%>

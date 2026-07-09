
<!-- #include file="database/dbconnect.asp" -->

<%

'''Generate PDF

Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument
Doc.ImportFromUrl ImportFromUrlpath & "rm_cn_new_print_pdf.asp?cn_no=" & request("cn_no")

Filename = Doc.Save(Server.MapPath(documentpath & request("cn_no") & ".pdf"), true )	
response.redirect ImportFromUrlpath & "/shared/" & request("cn_no") & ".pdf"
%>

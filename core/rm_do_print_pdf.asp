
<!-- #include file="database/datastore.asp" -->

<%

'''Generate PDF

Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument
Doc.ImportFromUrl ImportFromUrlpath & "rm_do_print.asp?do_code=" & request("do_code")

Filename = Doc.Save(Server.MapPath(documentpath & request("do_code") & ".pdf"), true )
documents = Server.MapPath(documentpath & request("do_code") & ".pdf")
	
'response.write "<a href='" & ImportFromUrlpath & "/doc/" & request("do_code") & ".pdf'>Click Here to view " & request("do_code") & "</a>"

response.redirect ImportFromUrlpath & "/doc/" & request("do_code") & ".pdf"
%>
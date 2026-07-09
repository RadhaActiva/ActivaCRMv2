<%

Set Mail = Server.CreateObject("Persits.MailSender")
Mail.Host = ""
Mail.From = ""
Mail.FromName = ""
Mail.AddAddress ""
Mail.Username = "" 
Mail.Password = "" 
Mail.Subject = ""
Mail.body = ""
Mail.port = 587
Mail.IsHTML = true
Mail.Send
Set Mail = nothing
response.write "sent"
%>
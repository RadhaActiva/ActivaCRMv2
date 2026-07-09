<!--#include file="dbconnect.local.asp"-->
<%
Dim strconnect

function CUD(sql)
set con = server.CreateObject("adodb.connection")
con.open strconnect
con.Execute sql
con.Close 
set con = nothing
CUD = "done"
end function

function selectid(sql)
set rs = server.CreateObject("adodb.recordset")
rs.Open sql,strconnect,0,1
if not rs.eof then
selectid = rs.Fields(0)
end if
rs.Close
set rs = nothing
end function

Function ChkString(str)
     if not isnull(str) and str <> "" then
	    str = Replace(str, "'", "`")
		str = Replace(str, """", "`")
	 end if 	 
	 ChkString = str
End Function

Function ChkStringLogin(str)

     if isnull(str) then
	 else
		 str = Replace(str, "'", "")	
		 str = Replace(str, "|", "")	
		 str = Replace(str, "&", "")	
		 str = Replace(str, ";", "")	
		 str = Replace(str, "$", "")	
		 str = Replace(str, "%", "")	
		 str = Replace(str, "@", "")	
		 str = Replace(str, "/", "")	
		 str = Replace(str, "\", "")	
		 str = Replace(str, "<", "")	
		 str = Replace(str, ">", "")	
		 str = Replace(str, "(", "")	
		 str = Replace(str, ")", "")	
		 str = Replace(str, "+", "")	
		 str = Replace(str, ",", "")	 
	 end if
	 
	 ChkStringLogin = str
End Function

Function ChkNumber0(itm)
if itm <> "" or not isnull(itm) then
ChkNumber0 = formatnumber(itm,0)
else
ChkNumber0 = itm
end if
End Function

Function ChkNumber1(itm)
if isnumeric(itm) then
ChkNumber1 = formatnumber(itm,1)
else
ChkNumber1 = itm
end if
End Function

Function ChkNumber2(itm)
if isnumeric(itm) then
	itm = formatnumber(itm,2)
	itm = Replace(itm, ",", "")
	ChkNumber2 = itm
else
	ChkNumber2 = "0.00"
end if
End Function

Function ChkNumber3(itm)
if isnumeric(itm) then
   ChkNumber3 = round(itm,2)
else
   ChkNumber3 = "0.00"
end if
End Function

Function ChkNumber2Decimal(itm)
if isnumeric(itm) then
   ChkNumber2Decimal = formatnumber(itm,2)
else
   ChkNumber2Decimal = "0.00"
end if
End Function

Function convertmonth(smonth)
dim stmp
   Select Case smonth
      Case 1   stmp = "Jan"
      Case 2   stmp = "Feb"
      Case 3   stmp = "Mar"
      Case 4   stmp = "Apr"
      Case 5   stmp = "May"
      Case 6   stmp = "Jun"
      Case 7   stmp = "Jul"
      Case 8   stmp = "Aug"
      Case 9   stmp = "Sep"
      Case 10  stmp = "Oct"
      Case 11  stmp = "Nov"
      Case 12  stmp = "Dec"
   End Select
   convertmonth = stmp
end Function

Function ChkDate(str)	  
	 ChkDate = day(str) & "-" & convertmonth(month(str)) & "-" & year(str)
End Function

Function replaceTime(itm)
  itm = FormatDateTime(itm,3)
  replaceTime = Mid(itm,1,Instr(itm,":") + 2) & " " & Right(itm,2)
End Function

Function ChkDateTime(itm)
if itm <> "" then
ChkDateTime = ChkDate(itm) & " " & replaceTime(itm)
end if
End Function

Function ChkDateTimeMySQL(itm)
if itm <> "" then
ChkDateTimeMySQL = Year(itm) & "-" & month(itm) & "-" & day(itm) & " " & hour(itm) & ":" & minute(itm) & ":" & second(itm)
end if
End Function

Function LenDecrypt(userName,userPassword)
  Dim userPasswordLength
  Dim userNameLength
  Dim Index
  Dim Ascii
  Dim Key
  Dim NewAscii
  Dim Decrypted
  If userPassword = "" Or userName = "" Then 
   	LenDecrypt = userPassword
    Exit function
  End if
  userPasswordLength = Len(userPassword) \ 2
  userNameLength = Len(userName)
  For Index = 1 To userPasswordLength
  	Ascii = CInt("&h" & Mid(userPassword, ((Index -1) * 2) + 1, 2))
  	Key = Asc(Mid(userName, (Index Mod userNameLength + 1), 1))
  	NewAscii = Ascii - Key
  	If NewAscii < 0 Then NewAscii = NewsAscii + 255
    Decrypted = Decrypted & Chr(NewAscii)
  Next
  LenDecrypt = Decrypted
End function

Function LenEncrypt(userName,userPassword)
  Dim userPasswordLength
  Dim userNameLength
  Dim Index
  Dim Ascii
  Dim Key
  Dim NewAscii
  Dim Encrypted
  If userPassword = "" Or userName = "" Then 
   	LenEncrypt = userPassword
   	Exit Function
  End If
  userPasswordLength = Len(userPassword)
  userNameLength = Len(userName)
  For Index = 1 To userPasswordLength
  	Ascii = Asc(Mid(userPassword, Index, 1))
  	Key = Asc(Mid(userName, (Index Mod userNameLength + 1), 1))
  	NewAscii = Ascii + Key
  	If NewAscii > 255 Then NewAscii = NewsAscii - 255
    Encrypted = Encrypted & Right("0" & Hex(NewAscii), 2)
  Next
  LenEncrypt = Left(Encrypted, 24)
End Function

Function sendemail(sname,semail,ssubject,ssbody)
'Set Mail = Server.CreateObject("Persits.MailSender")
'Mail.Host = "0.0.0.0"
'Mail.From = ""
'Mail.FromName = sname
'Mail.AddAddress semail
'Mail.Username = "" 'Please specify the valid user account here
'Mail.Password = "" ' Please specify you actual user password here
'Mail.Subject = ssubject
'Mail.body = ssbody
'Mail.port = 2525
'Mail.IsHTML = true
'Mail.Send
Set Mail = nothing
end function

Function sendemailA(sname,semail,ssubject,ssbody,docs)
'Set Mail = Server.CreateObject("Persits.MailSender")
'Mail.Host = "0.0.0.0"
'Mail.From = sname
'Mail.FromName = sname
'Mail.AddAddress semail
'Mail.Username = "" 'Please specify the valid user account here
'Mail.Password = "" ' Please specify you actual user password here
'Mail.AddAttachment docs
'Mail.Subject = ssubject
'Mail.body = ssbody
'Mail.port = 2525
'Mail.IsHTML = true
'Mail.Send
Set Mail = nothing
end function
%>

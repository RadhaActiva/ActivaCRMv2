<!--#include file="datastore.local.asp"-->
<%
Response.Buffer = True
'on error resume next
if Request.Cookies("GAPS")("loginstatus") <> "yes" then
   response.clear
   response.redirect "../login.asp"
end if   

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

Function isNotExceed(mmodel, issueqty) '  101122 - returns false if issued qty exceeds available qty
	isNotExceed="False"
	if trim(mmodel) <> "" and isnumeric(issueqty)  then
		sqlqty = "select sum(tblstocktran.stk_qty) AS totqty from tblstocktran inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code where tblmodel.md_code ='" & mmodel & "'"
  		maxqty=selectid(sqlqty)
		if not isnull(maxqty) and cint(issueqty) >= 1 then ' issue qty must be 1 or more and totalqty cannot be null
			if cint(maxqty) >= cint(issueqty) then
			  isNotExceed = "True"
			else
			  isNotExceed = "False"
			end if
		end if
	end if
	
	sql = "select md_category from tblmodel where md_code = '" & mmodel & "'"
	part_category = selectid(sql)
			
	if part_category="Service"  then '120325 skip checking / auto assign if these are not parts
			  isNotExceed = "True"
	end if	
End Function

Function InvoiceNumbering(str) '0509204  format follows 000001
     if not isnull(str) and str <> "" then
	    if len(str) = 1 then
	       str = "00000" & str
		elseif len(str) = 2 then
			 str = "0000" & str
	    elseif len(str) = 3 then
			str = "000" & str
	    elseif len(str) = 4 then
			str = "00" & str
	    elseif len(str) = 5 then
			str = "0" & str
		end if
	 end if 	
	 InvoiceNumbering = str
End Function

Function DONumbering(str) '0509204  format follows 000001
     if not isnull(str) and str <> "" then
	    if len(str) = 1 then
	       str = "00000" & str
		elseif len(str) = 2 then
			 str = "0000" & str
	    elseif len(str) = 3 then
			str = "000" & str
	    elseif len(str) = 4 then
			str = "00" & str
	    elseif len(str) = 5 then
			str = "0" & str
		end if
	 end if 	
	 DONumbering = str
End Function

Function ChkString(str)
     if not isnull(str) and str <> "" then
	    str = Replace(str, "'", "`")
		str = Replace(str, """", "`")
	 end if 	 
	 ChkString = str
End Function

Function ChkStringLogin(str)

     if isnull(str) or str = "" then
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

Function ChkNumber(str)
if isnumeric(str) then       
	ChkNumber = round(ccur(str),2)
elseif isnull(str) then
	ChkNumber = 0
else
	ChkNumber = round(ccur(str),2)
end if	 
End Function

Function ChkNumberInt(str)
if not isnumeric(str) then
	ChkNumberInt = "0"
else
	ChkNumberInt = str
end if
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

Function ChkNumber2(itm) 'updated 12022026
    If IsNull(itm) Or Trim(itm & "") = "" Then
        ChkNumber2 = "0.00"
        Exit Function
    End If

    ' remove thousand separators if the value is already formatted
    itm = Replace(itm, ",", "")

    If IsNumeric(itm) Then
        ChkNumber2 = FormatNumber(CDbl(itm), 2, -1, 0, 0)
    Else
        ChkNumber2 = "0.00"
    End If
End Function


'Function ChkNumber2(itm)
'if isnumeric(itm) then
'	itm = formatnumber(itm,2)
'	itm = Replace(itm, ",", "")
'	ChkNumber2 = itm
'else
'	ChkNumber2 = "0.00"
'end if
'End Function

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
     if isdate(str) then 
	    ChkDate = day(str) & "-" & convertmonth(month(str)) & "-" & year(str)
	 else
	    ChkDate = "" 	
	 end if
End Function

Function ChkDateYYYYMMDD(str)	  
     if isdate(str) then 
	 
	    if len(day(str)) = 1 then
		   sday = "0" & day(str)
		else
		   sday = day(str)
		end if
		
		if len(month(str)) = 1 then
		   smonth = "0" & month(str)
		else
		   smonth = month(str)    
		end if
		
	    ChkDateYYYYMMDD = year(str) & "-" & smonth & "-" & sday
	 else
	    ChkDateYYYYMMDD = "" 	
	 end if
End Function

Function ChkDateDDMMYYYY(str)	  
     if str <> "" and len(str) = 10 then 
	    ChkDateDDMMYYYY = left(str,2) & "-" & convertmonth(mid(str,4,2)) & "-" & right(str,4)
	 else
	    ChkDateDDMMYYYY = "" 	
	 end if
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

Function DateDiffDay(str1,str2)
	DateDiffDay = DateDiff("d",str1,str2)
End Function	


Function FilenameChange(filestr,xtension) 'renaming a file by generating random number for files that exists
	posisi= InStr(filestr, ".")
	filenameonly= trim(left(filestr,posisi-1))
	FilenameChange=filenameonly & "_" & Minute (Now()) & Second(Now()) & "." & xtension
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
If Len(MailHost & "") > 0 Then
	Set Mail = Server.CreateObject("Persits.MailSender")
	Mail.Host = MailHost
	Mail.From = MailFrom
	Mail.FromName = sname
	Mail.AddAddress semail
	Mail.Username = MailUsername
	Mail.Password = MailPassword
	Mail.Subject = ssubject
	Mail.body = ssbody
	Mail.port = MailPort
	Mail.IsHTML = true
	Mail.Send
	Set Mail = nothing
End If
end function

Function sendemailA(sname,semail,ssubject,ssbody,docs)
If Len(MailHost & "") > 0 Then
	Set Mail = Server.CreateObject("Persits.MailSender")
	Mail.Host = MailHost
	Mail.From = sname
	Mail.FromName = sname
	Mail.AddAddress semail
	Mail.Username = MailUsername
	Mail.Password = MailPassword
	Mail.AddAttachment docs
	Mail.Subject = ssubject
	Mail.body = ssbody
	Mail.port = MailPort
	Mail.IsHTML = true
	Mail.Send
	Set Mail = nothing
End If
end function

'16/10/2024 used for leading zeros for jobsheet/invoice/do numbers
'intValue is the number you want to pad
'intLength is the length you want to pad it out to
'Returned as a string either way
Function ZeroPadLeft(intValue, intLength)
    If Len(intValue) < intLength Then
       ZeroPadLeft = Right(Replace(Space(intLength), " ", "0") & intValue, intLength)
    Else
        ZeroPadLeft = CStr(intValue)
    End If
End Function
%>

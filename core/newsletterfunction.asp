<%
Function fixQuote(itm,cat)
  Select Case cat
    Case 0
	  If Not itm = "" Then
  	  fixQuote = Trim(Replace(itm,"'","''"))	 
	  fixQuote = Trim(Replace(fixQuote,chr(8),""))
  	  fixQuote = Trim(Replace(fixQuote,chr(9),""))
      fixQuote = Trim(Replace(fixQuote,chr(10),"")) 
	  fixQuote = Trim(Replace(fixQuote,"иC","-"))
	  fixQuote = Trim(Replace(fixQuote,"б▒",chr(34)))
	  fixQuote = Trim(Replace(fixQuote,"б░",chr(34)))
	  fixQuote = Trim(Replace(fixQuote,"бо","''"))
	  fixQuote = Trim(Replace(fixQuote,"бп","''"))
	  fixQuote = Trim(Replace(fixQuote,"С","''"))
	  fixQuote = Trim(Replace(fixQuote,"Е","..."))
	  fixQuote = Trim(Replace(fixQuote,"Т","''"))
	  fixQuote = Trim(Replace(fixQuote,"ик","&iacute;"))
	  End If
	Case 1
	  If Not IsNumeric(itm) Then
	    Response.End()
		Response.Redirect("index.htm")
	  Else
	    fixQuote = itm
	  End If
	Case 2
	  If Not itm = "" Then
  	  fixQuote = Trim(Replace(itm,chr(34),"&#34;"))
	  fixQuote = Trim(Replace(fixQuote,chr(8),""))
  	  fixQuote = Trim(Replace(fixQuote,chr(9),""))
      fixQuote = Trim(Replace(fixQuote,chr(10),"")) 
	  fixQuote = Trim(Replace(fixQuote,"иC","-"))
	  fixQuote = Trim(Replace(fixQuote,"б▒",chr(34)))
	  fixQuote = Trim(Replace(fixQuote,"б░",chr(34)))
	  fixQuote = Trim(Replace(fixQuote,"бо","'"))
	  fixQuote = Trim(Replace(fixQuote,"бп","'"))
	  fixQuote = Trim(Replace(fixQuote,"С","'"))
	  fixQuote = Trim(Replace(fixQuote,"Е","..."))
	  fixQuote = Trim(Replace(fixQuote,"Т","'"))
	  fixQuote = Trim(Replace(fixQuote,"ик","&iacute;"))
	  End If
	Case 3
	  If Not itm = "" Then
	  fixQuote = Trim(Replace(itm,"иC","-"))
	  fixQuote = Trim(Replace(fixQuote,"б▒",chr(34)))
	  fixQuote = Trim(Replace(fixQuote,"б░",chr(34)))
	  fixQuote = Trim(Replace(fixQuote,"бо","'"))
	  fixQuote = Trim(Replace(fixQuote,"бп","'"))
	  fixQuote = Trim(Replace(fixQuote,"ик","&iacute;"))
	  End If
  End Select
End Function

Function chgBreak(itm)
  If Not itm = "" Then
    chgBreak = Trim(Replace(itm,chr(13),"<br>"))
  End If
End Function

Function chgDate(itm)
  If Not itm = "" Then
    itm = Split(itm,"-")
    chgDate = itm(1) & "/" & itm(0) & "/" & itm(2)
  End If
End Function

Function fixDate(itm)
  fixDate = "'" & itm & "'"
End Function

Sub delFile(itm)
  itm = Server.MapPath(itm)
  Set FSO = Server.CreateObject( "Scripting.FileSystemObject" )
  If FSO.FileExists(itm) Then FSO.DeleteFile(itm)
  Set FSO = Nothing
End Sub

Sub createAFolder(itm)
  itm = Server.MapPath(itm)
  Set FSO = Server.CreateObject("Scripting.FileSystemObject")
  If Not FSO.FolderExists(itm) Then FSO.CreateFolder(itm)
  Set FSO = Nothing
End Sub

Sub copyAFolder(itm1,itm2)
  itm1 = Server.MapPath(itm1)
  itm2 = Server.MapPath(itm2)
  Set FSO = Server.CreateObject("Scripting.FileSystemObject")
  If Not FSO.FolderExists(itm2) Then FSO.CreateFolder(itm2)
  FSO.CopyFolder itm1, itm2
  Set FSO = Nothing
End Sub

Sub delFolder(itm)
  itm = Server.MapPath(itm)
  Set FSO = Server.CreateObject("Scripting.FileSystemObject")
  If FSO.FolderExists(itm) Then FSO.DeleteFolder(itm)
  Set FSO = Nothing
End Sub

Sub updateLog(logTable,logAction)

End Sub

Function chkLang(Eitm,Citm)
  If Session("lang") = "C" Then
    chkLang = Citm
  Else
    chkLang = Eitm
  End If
End Function

Function chkRadio(itm1,itm2)
  If itm1 = itm2 Then chkRadio = " CHECKED"
End Function

Function chkSelect(itm1,itm2)
  If itm1 = itm2 Then chkSelect = " SELECTED"
End Function
%>
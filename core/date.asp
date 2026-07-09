<%Option Explicit
Response.Buffer = true
Response.Flush 
%>
<html>
<head>
<title>Calender</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
a {text-decoration:none}
</style>
<script language="javascript">
function submitinfo() {window.document.frmcal.submit();}
function displayDate(sDay, sMonth, sYear) {	
//var month_of_year = new Array('January','February','March','April','May','June','July','August','September','October','November','December');
var month_of_year = new Array('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec');
var cmonth;    
    cmonth = month_of_year[sMonth-1];
	this.dateField = opener.dateField;
	this.inDate = dateField.value;	
	dateField.value = sDay + "-" + cmonth + "-" + sYear;	
	window.close()
}
</script>
</head>
<body bgcolor="#CCCCCC" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" alink="#000000" vlink="#000000" link="#000000">
<table cellpadding="3" cellspacing="0" width="100" border="1" bgcolor="#E6E6E6" align="center">
  <tr valign="middle" align="center">
    <form name="frmcal" action="date.asp" method="post">
    <td colspan="7" bgcolor="#999999">
      <select name="mth" onchange="submitinfo()">
<%
Dim dtdate, imonth, iyear, currmonth, currday, i, j
dtdate = Request.QueryString ("dt")
If dtdate <> "" Then
  imonth = Month(dtdate)
  iyear = Year(dtdate)
Else
  imonth = Request.Form("mth")
  iyear = Request.Form("yr")
  If imonth = "" Then imonth = Month(Date)
  If iyear = "" Then iyear = Year(Date)
End If
currmonth = DateSerial(iyear,imonth,1)

For i = 1 To 12
  If CInt(imonth) = i Then
    Response.Write "<option value='" & i & "' selected>" & MonthName(i) & "</option>"
  Else
    Response.Write "<option value='" & i & "'>" & MonthName(i) & "</option>"
  End If
Next

Response.Write "</select><select name='yr' onchange='submitinfo()'>"

For i = Year(Date) - 60 To Year(Date) + 3
  If CInt(iyear) = i Then
    Response.Write "<option value='" & i & "' selected>" & i & "</option>"
  Else
    Response.Write "<option value='" & i & "'>" & i & "</option>"
  End If
Next
%>
      </select> 
    </td>
    </form>
  </tr>
  <tr valign="middle" align="center">
    <td colspan="7">
<%
Response.Write _
      "<font face='Arial' color='#30679F'><b><a href='date.asp?dt=" & DateAdd("m",-1,currmonth) & "'><</a>" & _
      "<small> " & MonthName(imonth) & " " & iyear & " </small>" & _
      "<a href='date.asp?dt=" & DateAdd("m",1,currmonth) & "'>></a></b></font>" & _
    "</td>" & _
  "</tr>" & _
  "<tr valign='top' align='center' bgcolor='#000099'>"

For i = vbSunday To vbSaturday
  Response.Write "<th width='14%' bgcolor='#30679F'><font face='Arial' size='-2' color='#FFFFFF'>" & Left(WeekDayName(i),3) & "</font></th>"
Next
Response.Write "</tr>"

currday = currmonth
Do While WeekDay(currday) > vbSunday
  currday = DateAdd("d", -1, CDate(currday))
Loop

For i = 0 To 5
  Response.Write "<tr valign='top'>"
  For j = 0 To 6
    If Month(currday) = Month(currmonth) Then
      If currday = Date Then
        Response.Write "<td height='20' bgcolor='#CFC8CF' align='center'>"
      Else
        Response.Write "<td height='20' bgcolor='#FFFFFF' align='center'>"
      End If 
    Else
      Response.Write "<td height='20' bgcolor='#E6E6E6'>"
    End If
    If Month(currday) = Month(currmonth) Then
      Response.Write "<font face='Arial' size='-2'><b>" & _
        "<a href='javascript:displayDate(" & Day(currday) & "," & Month(currday) & "," & Year(currday) & ")'>"  & Day(currday) & "</a></b></font>"
    End If
    Response.Write "</td>"
    currday = DateAdd("d", 1, currday)
  Next
  Response.Write "</tr>"
Next
%>
</table>
</body>
</html>
<html></html>
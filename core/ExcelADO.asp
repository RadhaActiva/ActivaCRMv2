<%
' ExcelADO 1.0
' Excel File access via ADO
' Written by Hunter Beanland
' http://www.geocities.com/hbeanland/
' hbeanland@yahoo.com.au

Class ExcelADO
	Public ErrorText
	Public VirtualFileName
	Public AbsoluteFileName
	Public Tables()
	Public Columns()
	Public NumTables
	Public NumColumns
	Public SQLStatement
	Public objRSXL
	Public forReading
	Public forWriting
	Public Driver
	Public DriverJet
	Public DriverExcel
	Public DriverText
	Private Row
	Private objExcel
	
	Private Sub Class_Initialize
		ErrorText = ""
		SQLStatement = ""
		forReading = 0
		forWriting = 1
		DriverJet = 0
		DriverExcel = 1
		DriverText = 2
		Driver = 0
	End Sub
	
	Private Function GetWorksheets
	Dim objSchema
		set objSchema = objExcel.OpenSchema(20)
		objSchema.MoveFirst
		NumTables =0
		'objSchema.GetRows brings back 2 dimensions not one.
		while not objSchema.EOF
			Redim preserve Tables(NumTables)
			Tables(NumTables) = objSchema("TABLE_NAME")
			objSchema.MoveNext
			NumTables = NumTables+1
		wend
		if NumTables =0 then
			ErrorText = "No named areas or worksheet tab names found"
			GetWorksheets = False
			exit function
		end if
		objSchema.Close
		GetWorksheets = True
	End Function
	
	Private Sub GetColumns
		Dim objXLField
		NumColumns =0
		'.GetRows brings back 2 dimensions not one.
		For Each objXLField in objRSXL.Fields
			Redim preserve Columns(NumColumns)
			Columns(NumColumns) = objXLField.Name
			NumColumns = NumColumns + 1
		Next
		if NumColumns =0 then
			ErrorText = "No named areas or column names found"
		end if
	End Sub

	Public Function OpenSpreadsheet(FileName)
		Set objExcel = CreateObject("ADODB.Connection")
		Set objRSXL = CreateObject("ADODB.Recordset")
		ErrorText = ""
		'Check and fix file name
		if instr(trim(FileName),":") = 2 then 
			AbsoluteFileName = trim(FileName)
		elseif instr(trim(FileName),".") = 1 then
			ErrorText = "Path must be full virtual or absolute (. and .. not allowed)"
			OpenSpreadsheet = False
			exit Function
		else
			VirtualFileName = trim(FileName)
			AbsoluteFileName = Server.MapPath(trim(FileName))  
			'AbsoluteFileName = "e:\User\mothercare.com.my\web\web\propic\VIPTextFile.xls" 
		end if
		'Default to Microsoft Jet driver
		'http://support.microsoft.com/default.aspx?scid=kb;EN-US;257819
		if Driver = DriverExcel then
			objExcel.ConnectionString = "Provider=MSDASQL;Driver={Microsoft Excel Driver (*.xls)}; DBQ=" & AbsoluteFileName & ";ReadOnly=0;"
			'Other options: DriverID=278;DefaultDir=
		elseif Driver = DriverText then
			objExcel.ConnectionString = "Provider=MSDASQL;Driver={Microsoft Text Driver (*.txt; *.csv)}; DBQ=" & AbsoluteFileName & ";ReadOnly=0;"
		else	
			objExcel.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & AbsoluteFileName & "; Extended Properties=Excel 8.0;"
			'HDR=No - not supported in >= MDAC 2.6; IMEX=1 - Import Mixed; MaxScanRows=16
		end if	
		objExcel.CursorLocation = 3 'adUseClient
		On Error Resume Next
		objExcel.Open
		If objExcel.State = 0 then
			objExcel.Close
			ErrorText = "Connection will not open - check file path, lock status and valid Excel file"
			OpenSpreadsheet = False
		Else
			if GetWorksheets then
				OpenSpreadsheet = True
				ErrorText = ""
			else
				OpenSpreadsheet = False
			end if	
		End if
		On Error Goto 0
	End Function
	
	Public Function OpenWorksheet(TableName, IOMode)
		If objRSXL.State <> 0 then CloseWorksheet	'close existing worksheets
		ErrorText = ""
		if instr(join(Tables),TableName) = 0 then
			ErrorText = "Not valid named area or worksheet tab name"
			OpenWorksheet = False
			exit Function
		end if
		On Error Resume Next
		SQLStatement = "SELECT * FROM [" & TableName & "] " & SQLStatement
		if IOMode = forReading then
			objRSXL.Open SQLStatement, objExcel, 0, 1  'adOpenForwardOnly, adLockReadOnly
		else
			objRSXL.Open SQLStatement, objExcel, 2, 3  'adOpenDynamic, adLockOptimistic
		end if
		SQLStatement = ""	'reset SQL query so it is not carried forward to the next worksheet opening
		If objRSXL.State = 0 then
			objRSXL.Close
			ErrorText = "Worksheet will not open - check named area or worksheet tab name is defined"
			OpenWorksheet = False
		Else
			GetColumns
			OpenWorksheet = True
			ErrorText = ""
		End if
		On Error Goto 0
	End Function
	
	Public Function ReadRow
		'Read a row into the array
		if not objRSXL.EOF then
			Row = objRSXL.GetRows(1)
			ReadRow = true
		else
			ReadRow = false
		end if	
	End Function
	
	Public Sub CloseWorksheet
		'Close Recordset
		objRSXL.Close
	End Sub
	
	Public Function ColumnData(ColNum)
		if ColNum <= NumColumns-1 then
			ColumnData = Row(ColNum,0)
		else
			ColumnData = NULL
		end if	
	End Function
	
	Public Sub CloseSpreadsheet
		'Close and destroy
		On Error Resume Next
		If isobject(objRSXL) then
			If objRSXL.State <> 0 then objRSXL.Close
			set objRSXL = nothing
		end if	
		If isobject(objExcel) then
			If objExcel.State <> 0 then objExcel.Close
			set objExcel = nothing
		end if	
		On Error Goto 0
	End Sub
	
	Private Sub Class_Terminate
		'Make sure it is closed (to release file locks and resources)
		CloseSpreadsheet
	End Sub
	
End Class	

%>
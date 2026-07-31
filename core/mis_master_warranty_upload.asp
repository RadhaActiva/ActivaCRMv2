<%
If LCase(Trim(Request.Cookies("GAPS")("slevel") & "")) <> "mis" Then
    Response.Redirect "../login.asp"
End If
%>
<!-- #include file="header.asp" -->
<%
Dim warrantyMessage, warrantyMessageColor, currentWarrantyCount
warrantyMessage = ""
warrantyMessageColor = "#CC0000"

Function WarrantyHtml(value)
    WarrantyHtml = Server.HTMLEncode(CStr(value & ""))
End Function

Function WarrantyCleanText(value)
    Dim cleaned
    If IsNull(value) Or IsEmpty(value) Then
        WarrantyCleanText = ""
        Exit Function
    End If

    cleaned = CStr(value)
    cleaned = Replace(cleaned, Chr(9), " ")
    cleaned = Replace(cleaned, Chr(10), " ")
    cleaned = Replace(cleaned, Chr(13), " ")
    cleaned = Replace(cleaned, ChrW(&HFEFF), "")
    WarrantyCleanText = Trim(cleaned)
End Function

Function WarrantyOpenExcelConnection(filePath, fileExtension, ByRef connectionError)
    Dim excelConnection, excelProperties
    Set excelConnection = Server.CreateObject("ADODB.Connection")

    If fileExtension = "xlsx" Then
        excelProperties = "Excel 12.0 Xml;HDR=NO;IMEX=1"
    Else
        excelProperties = "Excel 8.0;HDR=NO;IMEX=1"
    End If

    On Error Resume Next
    excelConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & _
        ";Extended Properties=""" & excelProperties & """;"

    If Err.Number <> 0 Or excelConnection.State = 0 Then
        connectionError = Err.Description
        Err.Clear
        excelConnection.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & filePath & _
            ";Extended Properties=""" & excelProperties & """;"
    End If

    If Err.Number <> 0 Or excelConnection.State = 0 Then
        connectionError = Err.Description
        Err.Clear
        Set excelConnection = Nothing
    Else
        connectionError = ""
    End If
    On Error GoTo 0

    Set WarrantyOpenExcelConnection = excelConnection
End Function

Function WarrantyFirstWorksheet(excelConnection)
    Dim schemaRecordset, tableName, unquotedName
    WarrantyFirstWorksheet = ""

    Set schemaRecordset = excelConnection.OpenSchema(20)
    Do While Not schemaRecordset.EOF
        tableName = CStr(schemaRecordset("TABLE_NAME") & "")
        unquotedName = Replace(tableName, "'", "")
        If Right(unquotedName, 1) = "$" And InStr(1, tableName, "_FilterDatabase", 1) = 0 Then
            WarrantyFirstWorksheet = tableName
            Exit Do
        End If
        schemaRecordset.MoveNext
    Loop
    schemaRecordset.Close
    Set schemaRecordset = Nothing
End Function

Sub WarrantyDeleteFile(fileSystem, filePath)
    On Error Resume Next
    If Len(filePath) > 0 Then
        If fileSystem.FileExists(filePath) Then fileSystem.DeleteFile filePath, True
    End If
    On Error GoTo 0
End Sub

Function WarrantyHeaderAttribute(headers, attributeName)
    Dim marker, valueStart, valueEnd
    WarrantyHeaderAttribute = ""
    marker = attributeName & "=" & Chr(34)
    valueStart = InStr(1, headers, marker, 1)
    If valueStart = 0 Then Exit Function

    valueStart = valueStart + Len(marker)
    valueEnd = InStr(valueStart, headers, Chr(34), 0)
    If valueEnd > valueStart Then
        WarrantyHeaderAttribute = Mid(headers, valueStart, valueEnd - valueStart)
    End If
End Function

Function WarrantyReadCsvRows(filePath, ByRef csvError)
    Dim csvRows, csvStream, csvText, csvPosition, csvLength, currentCharacter
    Dim fieldValue, fieldIndex, rowIndex, rowValues(2), inQuotes

    Set csvRows = Server.CreateObject("Scripting.Dictionary")
    Set csvStream = Nothing
    csvError = ""

    On Error Resume Next
    Set csvStream = Server.CreateObject("ADODB.Stream")
    csvStream.Type = 2
    csvStream.Charset = "utf-8"
    csvStream.Open
    csvStream.LoadFromFile filePath
    csvText = csvStream.ReadText
    If Err.Number <> 0 Then
        csvError = "The CSV file could not be read. Save it as CSV UTF-8 and try again. " & Err.Description
        Err.Clear
    End If
    If Not csvStream Is Nothing Then
        If csvStream.State <> 0 Then csvStream.Close
    End If
    Set csvStream = Nothing
    On Error GoTo 0

    If csvError <> "" Then
        Set WarrantyReadCsvRows = csvRows
        Exit Function
    End If

    fieldValue = ""
    fieldIndex = 0
    rowIndex = 0
    inQuotes = False
    rowValues(0) = ""
    rowValues(1) = ""
    rowValues(2) = ""
    csvPosition = 1
    csvLength = Len(csvText)

    Do While csvPosition <= csvLength
        currentCharacter = Mid(csvText, csvPosition, 1)

        If currentCharacter = Chr(34) Then
            If inQuotes And csvPosition < csvLength And Mid(csvText, csvPosition + 1, 1) = Chr(34) Then
                fieldValue = fieldValue & Chr(34)
                csvPosition = csvPosition + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf currentCharacter = "," And Not inQuotes Then
            If fieldIndex <= 2 Then rowValues(fieldIndex) = fieldValue
            fieldIndex = fieldIndex + 1
            fieldValue = ""
        ElseIf (currentCharacter = Chr(13) Or currentCharacter = Chr(10)) And Not inQuotes Then
            If fieldIndex <= 2 Then rowValues(fieldIndex) = fieldValue
            rowIndex = rowIndex + 1
            csvRows.Add CStr(rowIndex), Array(rowValues(0), rowValues(1), rowValues(2))

            fieldValue = ""
            fieldIndex = 0
            rowValues(0) = ""
            rowValues(1) = ""
            rowValues(2) = ""

            If currentCharacter = Chr(13) And csvPosition < csvLength And Mid(csvText, csvPosition + 1, 1) = Chr(10) Then
                csvPosition = csvPosition + 1
            End If
        Else
            fieldValue = fieldValue & currentCharacter
        End If

        csvPosition = csvPosition + 1
    Loop

    If inQuotes Then
        csvError = "The CSV file contains an unterminated quoted field."
    ElseIf fieldValue <> "" Or fieldIndex > 0 Then
        If fieldIndex <= 2 Then rowValues(fieldIndex) = fieldValue
        rowIndex = rowIndex + 1
        csvRows.Add CStr(rowIndex), Array(rowValues(0), rowValues(1), rowValues(2))
    End If

    Set WarrantyReadCsvRows = csvRows
End Function

Sub WarrantyCollectRow(sourceRowNumber, sourceItemCode, sourceDescription, sourceMonth, ByRef headerFound, validationErrors, warrantyRows)
    Dim cleanItemCode, cleanDescription, cleanMonth, cleanMonthValue, cleanDuplicateKey

    cleanItemCode = WarrantyCleanText(sourceItemCode)
    cleanDescription = WarrantyCleanText(sourceDescription)
    cleanMonth = WarrantyCleanText(sourceMonth)

    If Not headerFound Then
        If LCase(cleanItemCode) = "item code" And LCase(cleanDescription) = "description" And LCase(cleanMonth) = "month" Then
            headerFound = True
        End If
        Exit Sub
    End If

    If cleanItemCode = "" And cleanDescription = "" And cleanMonth = "" Then Exit Sub

    If cleanItemCode = "" Then
        validationErrors.Add CStr(sourceRowNumber) & "_code", "Row " & sourceRowNumber & ": Item Code is required."
    ElseIf Len(cleanItemCode) > 50 Then
        validationErrors.Add CStr(sourceRowNumber) & "_code", "Row " & sourceRowNumber & ": Item Code exceeds 50 characters."
    End If

    If Len(cleanDescription) > 300 Then
        validationErrors.Add CStr(sourceRowNumber) & "_description", "Row " & sourceRowNumber & ": Description exceeds 300 characters."
    End If

    cleanMonthValue = Null
    If cleanMonth <> "" Then
        If Not IsNumeric(cleanMonth) Then
            validationErrors.Add CStr(sourceRowNumber) & "_month", "Row " & sourceRowNumber & ": Month must be a whole number."
        ElseIf CDbl(cleanMonth) <> Fix(CDbl(cleanMonth)) Or CDbl(cleanMonth) < 0 Or CDbl(cleanMonth) > 2147483647 Then
            validationErrors.Add CStr(sourceRowNumber) & "_month", "Row " & sourceRowNumber & ": Month must be a whole number from 0 to 2147483647."
        Else
            cleanMonthValue = CLng(cleanMonth)
        End If
    End If

    If cleanItemCode <> "" And Len(cleanItemCode) <= 50 Then
        cleanDuplicateKey = UCase(cleanItemCode)
        If warrantyRows.Exists(cleanDuplicateKey) Then
            validationErrors.Add CStr(sourceRowNumber) & "_duplicate", "Row " & sourceRowNumber & ": duplicate Item Code '" & cleanItemCode & "'."
        Else
            warrantyRows.Add cleanDuplicateKey, Array(cleanItemCode, cleanDescription, cleanMonthValue)
        End If
    End If
End Sub

Function WarrantyReceiveUpload(fileSystem, ByRef savedPath, ByRef fileExtension, ByRef originalFileName, ByRef receiveError)
    Dim contentType, boundaryPosition, boundary, delimiter
    Dim requestBytes, binaryStream, textStream, outputStream, bodyText
    Dim searchPosition, delimiterPosition, partStart, headerEnd, nextDelimiter
    Dim partHeaders, fieldName, candidateFileName, candidateStart, candidateLength
    Dim fileCount, pathPosition, dotPosition, temporaryFolder, uploadFolder, uniqueFileName

    WarrantyReceiveUpload = False
    savedPath = ""
    fileExtension = ""
    originalFileName = ""
    receiveError = ""
    Set binaryStream = Nothing
    Set textStream = Nothing
    Set outputStream = Nothing
    Set temporaryFolder = Nothing

    If Request.TotalBytes <= 0 Then
        receiveError = "Please select exactly one Excel workbook or CSV file."
        Exit Function
    End If

    If Request.TotalBytes > 5242880 Then
        receiveError = "The uploaded file exceeds the 5 MB limit."
        Exit Function
    End If

    contentType = Request.ServerVariables("HTTP_CONTENT_TYPE")
    boundaryPosition = InStr(1, contentType, "boundary=", 1)
    If InStr(1, contentType, "multipart/form-data", 1) = 0 Or boundaryPosition = 0 Then
        receiveError = "The upload request is not valid multipart form data."
        Exit Function
    End If

    boundary = Trim(Mid(contentType, boundaryPosition + Len("boundary=")))
    If Left(boundary, 1) = Chr(34) And Right(boundary, 1) = Chr(34) Then
        boundary = Mid(boundary, 2, Len(boundary) - 2)
    End If
    delimiter = "--" & boundary

    On Error Resume Next
    requestBytes = Request.BinaryRead(Request.TotalBytes)
    If Err.Number <> 0 Then
        receiveError = "The upload request could not be read. " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Set binaryStream = Server.CreateObject("ADODB.Stream")
    binaryStream.Type = 1
    binaryStream.Open
    binaryStream.Write requestBytes

    Set textStream = Server.CreateObject("ADODB.Stream")
    textStream.Type = 1
    textStream.Open
    textStream.Write requestBytes
    textStream.Position = 0
    textStream.Type = 2
    textStream.Charset = "iso-8859-1"
    bodyText = textStream.ReadText

    If Err.Number <> 0 Then
        receiveError = "The uploaded file could not be decoded for processing. " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If receiveError = "" Then
        fileCount = 0
        candidateStart = 0
        candidateLength = 0
        searchPosition = 1

        Do
            delimiterPosition = InStr(searchPosition, bodyText, delimiter, 0)
            If delimiterPosition = 0 Then Exit Do

            partStart = delimiterPosition + Len(delimiter)
            If Mid(bodyText, partStart, 2) = "--" Then Exit Do
            If Mid(bodyText, partStart, 2) = vbCrLf Then partStart = partStart + 2

            headerEnd = InStr(partStart, bodyText, vbCrLf & vbCrLf, 0)
            If headerEnd = 0 Then Exit Do

            partHeaders = Mid(bodyText, partStart, headerEnd - partStart)
            fieldName = WarrantyHeaderAttribute(partHeaders, "name")
            candidateFileName = WarrantyHeaderAttribute(partHeaders, "filename")
            nextDelimiter = InStr(headerEnd + 4, bodyText, vbCrLf & delimiter, 0)
            If nextDelimiter = 0 Then Exit Do

            If LCase(fieldName) = "warranty_file" And Len(Trim(candidateFileName)) > 0 Then
                fileCount = fileCount + 1
                If fileCount = 1 Then
                    originalFileName = candidateFileName
                    candidateStart = headerEnd + 4
                    candidateLength = nextDelimiter - candidateStart
                End If
            End If

            searchPosition = nextDelimiter + 2
        Loop

        If fileCount <> 1 Or candidateStart = 0 Or candidateLength <= 0 Then
            receiveError = "Please select exactly one Excel workbook or CSV file."
        End If
    End If

    If receiveError = "" Then
        originalFileName = Replace(originalFileName, "/", "\")
        pathPosition = InStrRev(originalFileName, "\")
        If pathPosition > 0 Then originalFileName = Mid(originalFileName, pathPosition + 1)

        dotPosition = InStrRev(originalFileName, ".")
        If dotPosition = 0 Then
            receiveError = "Only .csv, .xlsx or .xls files are accepted."
        Else
            fileExtension = LCase(Mid(originalFileName, dotPosition + 1))
            If fileExtension <> "csv" And fileExtension <> "xlsx" And fileExtension <> "xls" Then
                receiveError = "Only .csv, .xlsx or .xls files are accepted."
            End If
        End If
    End If

    If receiveError = "" Then
        On Error Resume Next
        Set temporaryFolder = fileSystem.GetSpecialFolder(2)
        uploadFolder = temporaryFolder.Path
        If Err.Number <> 0 Then
            receiveError = "The server temporary folder is unavailable. " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If

    If receiveError = "" Then
        Randomize
        uniqueFileName = "warranty_" & Year(Now()) & Right("0" & Month(Now()), 2) & _
            Right("0" & Day(Now()), 2) & Right("0" & Hour(Now()), 2) & _
            Right("0" & Minute(Now()), 2) & Right("0" & Second(Now()), 2) & _
            "_" & CStr(Int(Rnd() * 1000000)) & "." & fileExtension
        savedPath = fileSystem.BuildPath(uploadFolder, uniqueFileName)

        On Error Resume Next
        binaryStream.Position = candidateStart - 1
        Set outputStream = Server.CreateObject("ADODB.Stream")
        outputStream.Type = 1
        outputStream.Open
        binaryStream.CopyTo outputStream, candidateLength
        outputStream.SaveToFile savedPath, 2
        If Err.Number <> 0 Then
            receiveError = "The uploaded file could not be saved for processing. " & Err.Description
            Err.Clear
            WarrantyDeleteFile fileSystem, savedPath
            savedPath = ""
        End If
        On Error GoTo 0
    End If

    On Error Resume Next
    If Not outputStream Is Nothing Then
        If outputStream.State <> 0 Then outputStream.Close
    End If
    If Not textStream Is Nothing Then
        If textStream.State <> 0 Then textStream.Close
    End If
    If Not binaryStream Is Nothing Then
        If binaryStream.State <> 0 Then binaryStream.Close
    End If
    Set outputStream = Nothing
    Set textStream = Nothing
    Set binaryStream = Nothing
    Set temporaryFolder = Nothing
    On Error GoTo 0

    If receiveError = "" Then
        WarrantyReceiveUpload = True
    End If
End Function

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim fileExtension, temporaryFilePath, originalFileName
    Dim fileSystem, uploadError
    Dim excelConnection, worksheetName, excelRecordset, excelConnectionError
    Dim csvRows, csvRowKey, csvRowData, csvReadError
    Dim rowNumber, headerFound, itemCode, itemDescription, monthText, monthValue
    Dim validationErrors, warrantyRows, warrantyData, duplicateKey
    Dim databaseConnection, insertCommand, databaseError, insertedCount
    Dim databaseTransactionStarted

    temporaryFilePath = ""
    uploadError = ""
    databaseTransactionStarted = False
    Set excelConnection = Nothing
    Set fileSystem = Server.CreateObject("Scripting.FileSystemObject")

    Call WarrantyReceiveUpload(fileSystem, temporaryFilePath, fileExtension, originalFileName, uploadError)

    If uploadError = "" And fileExtension <> "csv" Then
        Set excelConnection = WarrantyOpenExcelConnection(temporaryFilePath, fileExtension, excelConnectionError)
        If excelConnection Is Nothing Then
            uploadError = "The workbook could not be opened. Ensure Microsoft Access Database Engine is installed on the server. " & excelConnectionError
        End If
    End If

    If uploadError = "" And fileExtension <> "csv" Then
        worksheetName = WarrantyFirstWorksheet(excelConnection)
        If Len(worksheetName) = 0 Then
            uploadError = "The workbook does not contain a readable worksheet."
        End If
    End If

    If uploadError = "" Then
        Set validationErrors = Server.CreateObject("Scripting.Dictionary")
        Set warrantyRows = Server.CreateObject("Scripting.Dictionary")
        warrantyRows.CompareMode = 1
        rowNumber = 0
        headerFound = False

        If fileExtension = "csv" Then
            Set csvRows = WarrantyReadCsvRows(temporaryFilePath, csvReadError)
            If csvReadError <> "" Then uploadError = csvReadError

            If uploadError = "" Then
                For Each csvRowKey In csvRows.Keys
                    rowNumber = CLng(csvRowKey)
                    csvRowData = csvRows(csvRowKey)
                    Call WarrantyCollectRow(rowNumber, csvRowData(0), csvRowData(1), csvRowData(2), headerFound, validationErrors, warrantyRows)
                    If Not headerFound And rowNumber >= 20 Then Exit For
                Next
            End If
            Set csvRows = Nothing
        Else
            Set excelRecordset = Server.CreateObject("ADODB.Recordset")
            On Error Resume Next
            excelRecordset.Open "SELECT F1, F2, F3 FROM [" & Replace(worksheetName, "]", "]]") & "]", excelConnection, 0, 1
            If Err.Number <> 0 Then
                uploadError = "The first three worksheet columns could not be read. " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

            If uploadError = "" Then
                Do While Not excelRecordset.EOF
                    rowNumber = rowNumber + 1
                    Call WarrantyCollectRow(rowNumber, excelRecordset.Fields(0).Value, excelRecordset.Fields(1).Value, excelRecordset.Fields(2).Value, headerFound, validationErrors, warrantyRows)
                    If Not headerFound And rowNumber >= 20 Then Exit Do
                    excelRecordset.MoveNext
                Loop
            End If

            If excelRecordset.State <> 0 Then excelRecordset.Close
            Set excelRecordset = Nothing
        End If

        If uploadError = "" Then
            If Not headerFound Then
                uploadError = "The required headers Item Code, Description and Month were not found in the first 20 rows."
            ElseIf warrantyRows.Count = 0 Then
                uploadError = "The uploaded file does not contain any warranty data rows."
            ElseIf validationErrors.Count > 0 Then
                uploadError = "Validation failed. " & Join(validationErrors.Items, " ")
            End If
        End If
    End If

    If uploadError = "" Then
        Set databaseConnection = Server.CreateObject("ADODB.Connection")
        databaseError = ""
        insertedCount = 0

        On Error Resume Next
        databaseConnection.Open strconnect
        If Err.Number <> 0 Then
            databaseError = Err.Description
            Err.Clear
        Else
            databaseConnection.BeginTrans
            If Err.Number <> 0 Then
                databaseError = Err.Description
                Err.Clear
            Else
                databaseTransactionStarted = True
                databaseConnection.Execute "DELETE FROM tbl_warranty"
                If Err.Number <> 0 Then
                    databaseError = Err.Description
                    Err.Clear
                End If
            End If
        End If

        If databaseError = "" Then
            For Each duplicateKey In warrantyRows.Keys
                warrantyData = warrantyRows(duplicateKey)
                Set insertCommand = Server.CreateObject("ADODB.Command")
                Set insertCommand.ActiveConnection = databaseConnection
                insertCommand.CommandType = 1
                insertCommand.CommandText = "INSERT INTO tbl_warranty (md_code, md_desc, [month]) VALUES (?, ?, ?)"
                insertCommand.Parameters.Append insertCommand.CreateParameter("md_code", 200, 1, 50, warrantyData(0))

                If warrantyData(1) = "" Then
                    insertCommand.Parameters.Append insertCommand.CreateParameter("md_desc", 200, 1, 300, Null)
                Else
                    insertCommand.Parameters.Append insertCommand.CreateParameter("md_desc", 200, 1, 300, warrantyData(1))
                End If

                insertCommand.Parameters.Append insertCommand.CreateParameter("month", 3, 1, , warrantyData(2))
                insertCommand.Execute

                If Err.Number <> 0 Then
                    databaseError = Err.Description
                    Err.Clear
                    Set insertCommand = Nothing
                    Exit For
                End If

                insertedCount = insertedCount + 1
                Set insertCommand = Nothing
            Next
        End If

        If databaseError = "" Then
            databaseConnection.CommitTrans
            If Err.Number <> 0 Then
                databaseError = Err.Description
                Err.Clear
            Else
                databaseTransactionStarted = False
            End If
        ElseIf databaseTransactionStarted Then
            databaseConnection.RollbackTrans
            databaseTransactionStarted = False
        End If

        If databaseError <> "" Then
            If databaseTransactionStarted Then
                databaseConnection.RollbackTrans
                databaseTransactionStarted = False
            End If
            uploadError = "The database was not changed because the replacement failed. " & databaseError
        End If

        If databaseConnection.State <> 0 Then databaseConnection.Close
        Set databaseConnection = Nothing
        On Error GoTo 0
    End If

    If Not excelConnection Is Nothing Then
        If excelConnection.State <> 0 Then excelConnection.Close
        Set excelConnection = Nothing
    End If
    WarrantyDeleteFile fileSystem, temporaryFilePath

    If uploadError = "" Then
        warrantyMessage = CStr(insertedCount) & " warranty records were imported successfully. Earlier records were replaced."
        warrantyMessageColor = "#008000"
    Else
        warrantyMessage = uploadError
    End If

    Set fileSystem = Nothing
End If

currentWarrantyCount = selectid("SELECT COUNT(*) FROM tbl_warranty")
%>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td class="titleblue1">&nbsp;</td>
                </tr>
                <tr>
                  <td class="titleblue1"><div align="left"><font color="#CC0000">Warranty</font> Upload</div></td>
                </tr>
              </table></td>
            </tr>
            <tr>
              <td valign="top" bgcolor="#FFFFFF"><table width="98%" border="0" cellspacing="0" cellpadding="3">
                <tr>
                  <td width="46%" align="left" bgcolor="#CCCCCC" scope="row"><strong>Master Setting &gt; Warranty Upload</strong></td>
                  <td align="left" bgcolor="#CCCCCC">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" align="left" valign="top" class="bodycopy">
                    <p>This replaces all records in <strong>tbl_warranty</strong>. The existing records remain unchanged if validation or import fails.</p>
                    <p>Accepted formats: <strong>CSV UTF-8</strong>, <strong>.xlsx</strong> or <strong>.xls</strong>. CSV works without an Excel provider on the server.</p>
                    <p>For the live server, open the workbook in Excel and use <strong>Save As &gt; CSV UTF-8 (Comma delimited) (*.csv)</strong> before uploading.</p>
                    <p>Expected columns: <strong>Item Code</strong>, <strong>Description</strong>, <strong>Month</strong>. The header may appear within the first 20 rows.</p>
                    <p>Current warranty records: <strong><%=WarrantyHtml(currentWarrantyCount)%></strong></p>
                    <% If warrantyMessage <> "" Then %>
                    <p style="color:<%=warrantyMessageColor%>;"><strong><%=WarrantyHtml(warrantyMessage)%></strong></p>
                    <% End If %>
                    <form action="mis_master_warranty_upload.asp" method="post" enctype="multipart/form-data" name="warrantyUploadForm" id="warrantyUploadForm" onsubmit="return confirm('Replace all existing warranty records with this workbook?');">
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr>
                          <td><strong>Warranty file</strong></td>
                          <td><input name="warranty_file" type="file" id="warranty_file" accept=".csv,.xlsx,.xls" required="required" /></td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                          <td><input name="Submit" type="submit" class="button" value="Upload and Replace" /></td>
                        </tr>
                      </table>
                    </form>
                  </td>
                </tr>
              </table></td>
            </tr>
          </table></td>
        </tr>
<!-- #include file="footer.asp" -->

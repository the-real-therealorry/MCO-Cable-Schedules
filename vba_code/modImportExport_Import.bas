' ===============================================================================
' IMPORT FUNCTIONS - Add these to modImportExport module
' ===============================================================================

' ===============================================================================
' PUBLIC API - IMPORT FUNCTIONS
' ===============================================================================

'-------------------------------------------------------------------------------
' FUNCTION: ImportCablesFromCSV
' PURPOSE: Imports cables from CSV file
' PARAMETERS: strFilePath - Full path to CSV file
'            importMode - "APPEND", "REPLACE", or "MERGE"
' RETURNS: Dictionary with import results and statistics
'-------------------------------------------------------------------------------
Public Function ImportCablesFromCSV(strFilePath As String, importMode As String) As Object
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim txtFile As Object
    Dim strLine As String
    Dim arrFields() As String
    Dim strPlantID As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim results As Object
    Dim cablesImported As Long
    Dim cablesSkipped As Long
    Dim errors As Collection
    Dim lineNum As Long

    ' Initialize
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.OpenTextFile(strFilePath, 1) ' 1 = ForReading
    Set results = CreateObject("Scripting.Dictionary")
    Set errors = New Collection

    lineNum = 0
    cablesImported = 0
    cablesSkipped = 0

    ' Skip header line
    If Not txtFile.AtEndOfStream Then
        strLine = txtFile.ReadLine
        lineNum = lineNum + 1
    End If

    ' Read and import each line
    Do While Not txtFile.AtEndOfStream
        strLine = txtFile.ReadLine
        lineNum = lineNum + 1

        ' Skip empty lines
        If Trim(strLine) = "" Then GoTo NextLine

        ' Parse CSV line
        arrFields = ParseCSVLine(strLine)

        ' Validate field count (13 fields expected)
        If UBound(arrFields) < 12 Then
            errors.Add "Line " & lineNum & ": Invalid field count"
            cablesSkipped = cablesSkipped + 1
            GoTo NextLine
        End If

        ' Extract fields
        ' arrFields: Version, Plant, Scheduled, IDAttached, CableID, Source, Destination,
        '           CoreSize, EarthSize, CoreConfig, InsulationType, CableType, CableLength
        strPlantID = arrFields(1)

        ' Get appropriate worksheet and table
        Select Case strPlantID
            Case "WET_PLANT"
                Set ws = sht_WetPlant
                Set tbl = ws.ListObjects("tbl_WetPlantCables")
            Case "ORE_SORTER"
                Set ws = sht_OreSorter
                Set tbl = ws.ListObjects("tbl_OreSorterCables")
            Case "RETREATMENT"
                Set ws = sht_Retreatment
                Set tbl = ws.ListObjects("tbl_RetreatmentCables")
            Case Else
                errors.Add "Line " & lineNum & ": Unknown plant ID: " & strPlantID
                cablesSkipped = cablesSkipped + 1
                GoTo NextLine
        End Select

        ' Handle import mode
        Select Case UCase(importMode)
            Case "APPEND"
                ' Add new row
                Set newRow = tbl.ListRows.Add
                PopulateCableRow newRow, arrFields
                cablesImported = cablesImported + 1

            Case "MERGE"
                ' Find existing cable by ID, update or add
                Dim existingRow As ListRow
                Dim cableID As String
                cableID = arrFields(4)

                Set existingRow = FindCableByID(tbl, cableID)
                If existingRow Is Nothing Then
                    ' Not found, add new
                    Set newRow = tbl.ListRows.Add
                    PopulateCableRow newRow, arrFields
                    cablesImported = cablesImported + 1
                Else
                    ' Found, update existing
                    PopulateCableRow existingRow, arrFields
                    cablesImported = cablesImported + 1
                End If

            Case "REPLACE"
                ' This mode should clear tables first (handled in calling function)
                Set newRow = tbl.ListRows.Add
                PopulateCableRow newRow, arrFields
                cablesImported = cablesImported + 1

            Case Else
                errors.Add "Invalid import mode: " & importMode
                Exit Do
        End Select

NextLine:
    Loop

    ' Close file
    txtFile.Close

    ' Build results dictionary
    results.Add "Success", True
    results.Add "CablesImported", cablesImported
    results.Add "CablesSkipped", cablesSkipped
    results.Add "Errors", errors

    Set ImportCablesFromCSV = results

    ' Cleanup
    Set txtFile = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    ' Build error results
    Set results = CreateObject("Scripting.Dictionary")
    results.Add "Success", False
    results.Add "ErrorMessage", Err.Description
    results.Add "ErrorLine", lineNum

    Set ImportCablesFromCSV = results

    If Not txtFile Is Nothing Then txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing

    Debug.Print "Error in ImportCablesFromCSV: " & Err.Number & " - " & Err.Description
End Function

'-------------------------------------------------------------------------------
' FUNCTION: ImportEndpointsFromCSV
' PURPOSE: Imports endpoints from CSV file
' PARAMETERS: strFilePath - Full path to CSV file
'            importMode - "APPEND", "REPLACE", or "MERGE"
' RETURNS: Dictionary with import results and statistics
'-------------------------------------------------------------------------------
Public Function ImportEndpointsFromCSV(strFilePath As String, importMode As String) As Object
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim txtFile As Object
    Dim strLine As String
    Dim arrFields() As String
    Dim strPlantID As String
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim results As Object
    Dim endpointsImported As Long
    Dim endpointsSkipped As Long
    Dim errors As Collection
    Dim lineNum As Long
    Dim strTableName As String

    ' Initialize
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.OpenTextFile(strFilePath, 1)
    Set results = CreateObject("Scripting.Dictionary")
    Set errors = New Collection

    lineNum = 0
    endpointsImported = 0
    endpointsSkipped = 0

    ' Skip header line
    If Not txtFile.AtEndOfStream Then
        strLine = txtFile.ReadLine
        lineNum = lineNum + 1
    End If

    ' Read and import each line
    Do While Not txtFile.AtEndOfStream
        strLine = txtFile.ReadLine
        lineNum = lineNum + 1

        ' Skip empty lines
        If Trim(strLine) = "" Then GoTo NextLine

        ' Parse CSV line
        arrFields = ParseCSVLine(strLine)

        ' Validate field count (4 fields expected: Version, Plant, ShortName, Description)
        If UBound(arrFields) < 3 Then
            errors.Add "Line " & lineNum & ": Invalid field count"
            endpointsSkipped = endpointsSkipped + 1
            GoTo NextLine
        End If

        ' Extract plant ID
        strPlantID = arrFields(1)

        ' Get appropriate endpoint table
        Select Case strPlantID
            Case "WET_PLANT"
                strTableName = "tbl_WetPlantEndpoints"
            Case "ORE_SORTER"
                strTableName = "tbl_OreSorterEndpoints"
            Case "RETREATMENT"
                strTableName = "tbl_RetreatmentEndpoints"
            Case Else
                errors.Add "Line " & lineNum & ": Unknown plant ID: " & strPlantID
                endpointsSkipped = endpointsSkipped + 1
                GoTo NextLine
        End Select

        Set tbl = sht_Data.ListObjects(strTableName)

        ' Handle import mode
        Select Case UCase(importMode)
            Case "APPEND"
                ' Add new row
                Set newRow = tbl.ListRows.Add
                newRow.Range(1, 1).Value = arrFields(2) ' ShortName
                newRow.Range(1, 2).Value = arrFields(3) ' Description
                endpointsImported = endpointsImported + 1

            Case "MERGE"
                ' Find existing endpoint by ShortName, update or add
                Dim existingRow As ListRow
                Dim shortName As String
                shortName = arrFields(2)

                Set existingRow = FindEndpointByShortName(tbl, shortName)
                If existingRow Is Nothing Then
                    ' Not found, add new
                    Set newRow = tbl.ListRows.Add
                    newRow.Range(1, 1).Value = arrFields(2)
                    newRow.Range(1, 2).Value = arrFields(3)
                    endpointsImported = endpointsImported + 1
                Else
                    ' Found, update description
                    existingRow.Range(1, 2).Value = arrFields(3)
                    endpointsImported = endpointsImported + 1
                End If

            Case "REPLACE"
                ' Add new row (tables cleared in calling function)
                Set newRow = tbl.ListRows.Add
                newRow.Range(1, 1).Value = arrFields(2)
                newRow.Range(1, 2).Value = arrFields(3)
                endpointsImported = endpointsImported + 1
        End Select

NextLine:
    Loop

    ' Close file
    txtFile.Close

    ' Build results
    results.Add "Success", True
    results.Add "EndpointsImported", endpointsImported
    results.Add "EndpointsSkipped", endpointsSkipped
    results.Add "Errors", errors

    Set ImportEndpointsFromCSV = results

    ' Cleanup
    Set txtFile = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    Set results = CreateObject("Scripting.Dictionary")
    results.Add "Success", False
    results.Add "ErrorMessage", Err.Description
    results.Add "ErrorLine", lineNum

    Set ImportEndpointsFromCSV = results

    If Not txtFile Is Nothing Then txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing

    Debug.Print "Error in ImportEndpointsFromCSV: " & Err.Number & " - " & Err.Description
End Function

' ===============================================================================
' PRIVATE HELPER FUNCTIONS FOR IMPORT
' ===============================================================================

'-------------------------------------------------------------------------------
' FUNCTION: ParseCSVLine
' PURPOSE: Parses a CSV line handling quoted fields and embedded commas
' PARAMETERS: strLine - CSV line to parse
' RETURNS: Array of field values
'-------------------------------------------------------------------------------
Private Function ParseCSVLine(strLine As String) As String()
    Dim fields() As String
    Dim fieldCount As Long
    Dim i As Long
    Dim inQuotes As Boolean
    Dim currentField As String
    Dim char As String

    ReDim fields(0)
    fieldCount = 0
    inQuotes = False
    currentField = ""

    For i = 1 To Len(strLine)
        char = Mid(strLine, i, 1)

        If char = """" Then
            ' Toggle quote state
            inQuotes = Not inQuotes
        ElseIf char = "," And Not inQuotes Then
            ' Field separator
            ReDim Preserve fields(fieldCount)
            fields(fieldCount) = currentField
            fieldCount = fieldCount + 1
            currentField = ""
        Else
            ' Regular character
            currentField = currentField & char
        End If
    Next i

    ' Add last field
    ReDim Preserve fields(fieldCount)
    fields(fieldCount) = currentField

    ParseCSVLine = fields
End Function

'-------------------------------------------------------------------------------
' SUB: PopulateCableRow
' PURPOSE: Populates a cable table row from CSV field array
' PARAMETERS: row - ListRow object to populate
'            arrFields - Array of field values from CSV
'-------------------------------------------------------------------------------
Private Sub PopulateCableRow(row As ListRow, arrFields() As String)
    ' arrFields: Version, Plant, Scheduled, IDAttached, CableID, Source, Destination,
    '           CoreSize, EarthSize, CoreConfig, InsulationType, CableType, CableLength

    row.Range(1, 1).Value = CBool(arrFields(2))    ' Scheduled
    row.Range(1, 2).Value = CBool(arrFields(3))    ' IDAttached
    row.Range(1, 3).Value = arrFields(4)           ' CableID
    row.Range(1, 4).Value = arrFields(5)           ' Source
    row.Range(1, 5).Value = arrFields(6)           ' Destination
    row.Range(1, 6).Value = arrFields(7)           ' CoreSize
    row.Range(1, 7).Value = arrFields(8)           ' EarthSize
    row.Range(1, 8).Value = arrFields(9)           ' CoreConfig
    row.Range(1, 9).Value = arrFields(10)          ' InsulationType
    row.Range(1, 10).Value = arrFields(11)         ' CableType
    row.Range(1, 11).Value = arrFields(12)         ' CableLength
End Sub

'-------------------------------------------------------------------------------
' FUNCTION: FindCableByID
' PURPOSE: Finds a cable row by Cable ID
' PARAMETERS: tbl - ListObject table to search
'            cableID - Cable ID to find
' RETURNS: ListRow object or Nothing if not found
'-------------------------------------------------------------------------------
Private Function FindCableByID(tbl As ListObject, cableID As String) As ListRow
    Dim row As ListRow
    Dim i As Long

    For i = 1 To tbl.ListRows.Count
        Set row = tbl.ListRows(i)
        If row.Range(1, 3).Value = cableID Then
            Set FindCableByID = row
            Exit Function
        End If
    Next i

    Set FindCableByID = Nothing
End Function

'-------------------------------------------------------------------------------
' FUNCTION: FindEndpointByShortName
' PURPOSE: Finds an endpoint row by Short Name
' PARAMETERS: tbl - ListObject table to search
'            shortName - Short name to find
' RETURNS: ListRow object or Nothing if not found
'-------------------------------------------------------------------------------
Private Function FindEndpointByShortName(tbl As ListObject, shortName As String) As ListRow
    Dim row As ListRow
    Dim i As Long

    For i = 1 To tbl.ListRows.Count
        Set row = tbl.ListRows(i)
        If row.Range(1, 1).Value = shortName Then
            Set FindEndpointByShortName = row
            Exit Function
        End If
    Next i

    Set FindEndpointByShortName = Nothing
End Function

'-------------------------------------------------------------------------------
' SUB: ClearPlantCables
' PURPOSE: Clears all cables from specified plant (for REPLACE mode)
' PARAMETERS: strPlantID - Plant identifier
'-------------------------------------------------------------------------------
Public Sub ClearPlantCables(strPlantID As String)
    Dim tbl As ListObject
    Dim ws As Worksheet

    Select Case strPlantID
        Case "WET_PLANT"
            Set ws = sht_WetPlant
            Set tbl = ws.ListObjects("tbl_WetPlantCables")
        Case "ORE_SORTER"
            Set ws = sht_OreSorter
            Set tbl = ws.ListObjects("tbl_OreSorterCables")
        Case "RETREATMENT"
            Set ws = sht_Retreatment
            Set tbl = ws.ListObjects("tbl_RetreatmentCables")
        Case "ALL"
            ClearPlantCables "WET_PLANT"
            ClearPlantCables "ORE_SORTER"
            ClearPlantCables "RETREATMENT"
            Exit Sub
    End Select

    ' Delete all rows
    On Error Resume Next
    If tbl.ListRows.Count > 0 Then
        tbl.DataBodyRange.Delete
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' SUB: ClearPlantEndpoints
' PURPOSE: Clears all endpoints from specified plant (for REPLACE mode)
' PARAMETERS: strPlantID - Plant identifier
'-------------------------------------------------------------------------------
Public Sub ClearPlantEndpoints(strPlantID As String)
    Dim tbl As ListObject
    Dim strTableName As String

    Select Case strPlantID
        Case "WET_PLANT"
            strTableName = "tbl_WetPlantEndpoints"
        Case "ORE_SORTER"
            strTableName = "tbl_OreSorterEndpoints"
        Case "RETREATMENT"
            strTableName = "tbl_RetreatmentEndpoints"
        Case "ALL"
            ClearPlantEndpoints "WET_PLANT"
            ClearPlantEndpoints "ORE_SORTER"
            ClearPlantEndpoints "RETREATMENT"
            Exit Sub
    End Select

    Set tbl = sht_Data.ListObjects(strTableName)

    ' Delete all rows
    On Error Resume Next
    If tbl.ListRows.Count > 0 Then
        tbl.DataBodyRange.Delete
    End If
    On Error GoTo 0
End Sub

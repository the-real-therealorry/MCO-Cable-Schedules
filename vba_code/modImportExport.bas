Attribute VB_Name = "modImportExport"
' ===============================================================================
' MODULE: modImportExport
' AUTHOR: AI Assistant
' DATE: 2024-12
' PURPOSE: Cable and endpoint import/export functionality
'          Supports CSV and JSON formats with version compatibility
' ===============================================================================

Option Explicit

' Module version for compatibility tracking
Public Const MODULE_VERSION As String = "2024.12.1"

' ===============================================================================
' PUBLIC API - EXPORT FUNCTIONS
' ===============================================================================

'-------------------------------------------------------------------------------
' FUNCTION: ExportCablesToCSV
' PURPOSE: Exports cables from specified plant(s) to CSV file
' PARAMETERS: strPlantID - "WET_PLANT", "ORE_SORTER", "RETREATMENT", or "ALL"
'            strFilePath - Full path where CSV should be saved
' RETURNS: Boolean - True if successful, False otherwise
'-------------------------------------------------------------------------------
Public Function ExportCablesToCSV(strPlantID As String, strFilePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim txtFile As Object
    Dim arrPlants() As String
    Dim plantID As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim i As Long
    Dim strLine As String
    Dim totalRows As Long

    ' Initialize file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(strFilePath, True)

    ' Write CSV header
    txtFile.WriteLine "Version,Plant,Scheduled,IDAttached,CableID,Source,Destination,CoreSize,EarthSize,CoreConfig,InsulationType,CableType,CableLength"

    ' Determine which plants to export
    If strPlantID = "ALL" Then
        arrPlants = Split("WET_PLANT,ORE_SORTER,RETREATMENT", ",")
    Else
        ReDim arrPlants(0)
        arrPlants(0) = strPlantID
    End If

    ' Export each plant
    For Each plantID In arrPlants
        ' Get the worksheet and table for this plant
        Select Case CStr(plantID)
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
                ' Unknown plant, skip
                GoTo NextPlant
        End Select

        ' Export each cable row
        If tbl.ListRows.Count > 0 Then
            totalRows = tbl.ListRows.Count

            For i = 1 To totalRows
                Set row = tbl.ListRows(i)

                ' Build CSV line
                strLine = MODULE_VERSION & "," & _
                         CStr(plantID) & "," & _
                         CSVEscape(row.Range(1, 1).Value) & "," & _
                         CSVEscape(row.Range(1, 2).Value) & "," & _
                         CSVEscape(row.Range(1, 3).Value) & "," & _
                         CSVEscape(row.Range(1, 4).Value) & "," & _
                         CSVEscape(row.Range(1, 5).Value) & "," & _
                         CSVEscape(row.Range(1, 6).Value) & "," & _
                         CSVEscape(row.Range(1, 7).Value) & "," & _
                         CSVEscape(row.Range(1, 8).Value) & "," & _
                         CSVEscape(row.Range(1, 9).Value) & "," & _
                         CSVEscape(row.Range(1, 10).Value) & "," & _
                         CSVEscape(row.Range(1, 11).Value)

                txtFile.WriteLine strLine

                ' Update progress if needed
                If i Mod 10 = 0 Then
                    DoEvents ' Allow UI to update
                End If
            Next i
        End If

NextPlant:
    Next plantID

    ' Close file
    txtFile.Close

    ' Success
    ExportCablesToCSV = True

    ' Cleanup
    Set txtFile = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    ExportCablesToCSV = False
    If Not txtFile Is Nothing Then txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing

    MsgBox "Error exporting cables to CSV: " & Err.Description, vbCritical, "Export Error"
    Debug.Print "Error in ExportCablesToCSV: " & Err.Number & " - " & Err.Description
End Function

'-------------------------------------------------------------------------------
' FUNCTION: ExportEndpointsToCSV
' PURPOSE: Exports endpoints from specified plant(s) to CSV file
' PARAMETERS: strPlantID - "WET_PLANT", "ORE_SORTER", "RETREATMENT", or "ALL"
'            strFilePath - Full path where CSV should be saved
' RETURNS: Boolean - True if successful, False otherwise
'-------------------------------------------------------------------------------
Public Function ExportEndpointsToCSV(strPlantID As String, strFilePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim txtFile As Object
    Dim arrPlants() As String
    Dim plantID As Variant
    Dim tbl As ListObject
    Dim row As ListRow
    Dim i As Long
    Dim strLine As String
    Dim strTableName As String

    ' Initialize file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(strFilePath, True)

    ' Write CSV header
    txtFile.WriteLine "Version,Plant,ShortName,Description"

    ' Determine which plants to export
    If strPlantID = "ALL" Then
        arrPlants = Split("WET_PLANT,ORE_SORTER,RETREATMENT", ",")
    Else
        ReDim arrPlants(0)
        arrPlants(0) = strPlantID
    End If

    ' Export each plant's endpoints
    For Each plantID In arrPlants
        ' Get the table name for this plant
        Select Case CStr(plantID)
            Case "WET_PLANT"
                strTableName = "tbl_WetPlantEndpoints"
            Case "ORE_SORTER"
                strTableName = "tbl_OreSorterEndpoints"
            Case "RETREATMENT"
                strTableName = "tbl_RetreatmentEndpoints"
            Case Else
                GoTo NextPlant
        End Select

        ' Get table reference
        Set tbl = sht_Data.ListObjects(strTableName)

        ' Export each endpoint
        If tbl.ListRows.Count > 0 Then
            For i = 1 To tbl.ListRows.Count
                Set row = tbl.ListRows(i)

                ' Build CSV line
                strLine = MODULE_VERSION & "," & _
                         CStr(plantID) & "," & _
                         CSVEscape(row.Range(1, 1).Value) & "," & _
                         CSVEscape(row.Range(1, 2).Value)

                txtFile.WriteLine strLine
            Next i
        End If

NextPlant:
    Next plantID

    ' Close file
    txtFile.Close

    ' Success
    ExportEndpointsToCSV = True

    ' Cleanup
    Set txtFile = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    ExportEndpointsToCSV = False
    If Not txtFile Is Nothing Then txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing

    MsgBox "Error exporting endpoints to CSV: " & Err.Description, vbCritical, "Export Error"
    Debug.Print "Error in ExportEndpointsToCSV: " & Err.Number & " - " & Err.Description
End Function

'-------------------------------------------------------------------------------
' FUNCTION: ExportToJSON
' PURPOSE: Exports cables and endpoints to JSON file
' PARAMETERS: strPlantID - "WET_PLANT", "ORE_SORTER", "RETREATMENT", or "ALL"
'            strFilePath - Full path where JSON should be saved
' RETURNS: Boolean - True if successful, False otherwise
'-------------------------------------------------------------------------------
Public Function ExportToJSON(strPlantID As String, strFilePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim txtFile As Object
    Dim arrPlants() As String
    Dim plantID As Variant
    Dim jsonOutput As String
    Dim plantJSON As String
    Dim totalCables As Long
    Dim totalEndpoints As Long

    ' Initialize
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(strFilePath, True)

    ' Determine which plants to export
    If strPlantID = "ALL" Then
        arrPlants = Split("WET_PLANT,ORE_SORTER,RETREATMENT", ",")
    Else
        ReDim arrPlants(0)
        arrPlants(0) = strPlantID
    End If

    ' Build JSON header
    jsonOutput = "{" & vbCrLf
    jsonOutput = jsonOutput & "  ""version"": """ & MODULE_VERSION & """," & vbCrLf
    jsonOutput = jsonOutput & "  ""exportDate"": """ & Format(Now, "yyyy-mm-ddThh:nn:ss") & "Z""," & vbCrLf
    jsonOutput = jsonOutput & "  ""sourceFile"": """ & ThisWorkbook.Name & """," & vbCrLf
    jsonOutput = jsonOutput & "  ""plants"": {" & vbCrLf

    ' Export each plant
    For Each plantID In arrPlants
        plantJSON = BuildPlantJSON(CStr(plantID), totalCables, totalEndpoints)
        jsonOutput = jsonOutput & plantJSON

        ' Add comma if not last plant
        If plantID <> arrPlants(UBound(arrPlants)) Then
            jsonOutput = jsonOutput & ","
        End If
        jsonOutput = jsonOutput & vbCrLf
    Next plantID

    ' Close plants object
    jsonOutput = jsonOutput & "  }," & vbCrLf

    ' Add metadata
    jsonOutput = jsonOutput & "  ""metadata"": {" & vbCrLf
    jsonOutput = jsonOutput & "    ""totalCables"": " & totalCables & "," & vbCrLf
    jsonOutput = jsonOutput & "    ""totalEndpoints"": " & totalEndpoints & "," & vbCrLf
    jsonOutput = jsonOutput & "    ""exportType"": """ & IIf(strPlantID = "ALL", "ALL_PLANTS", strPlantID) & """" & vbCrLf
    jsonOutput = jsonOutput & "  }" & vbCrLf

    ' Close root object
    jsonOutput = jsonOutput & "}"

    ' Write to file
    txtFile.Write jsonOutput
    txtFile.Close

    ' Success
    ExportToJSON = True

    ' Cleanup
    Set txtFile = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    ExportToJSON = False
    If Not txtFile Is Nothing Then txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing

    MsgBox "Error exporting to JSON: " & Err.Description, vbCritical, "Export Error"
    Debug.Print "Error in ExportToJSON: " & Err.Number & " - " & Err.Description
End Function

' ===============================================================================
' PRIVATE HELPER FUNCTIONS
' ===============================================================================

'-------------------------------------------------------------------------------
' FUNCTION: CSVEscape
' PURPOSE: Escapes special characters for CSV format
' PARAMETERS: value - Value to escape
' RETURNS: String - Escaped value suitable for CSV
'-------------------------------------------------------------------------------
Private Function CSVEscape(value As Variant) As String
    Dim strValue As String

    ' Handle null/empty
    If IsNull(value) Or IsEmpty(value) Then
        CSVEscape = ""
        Exit Function
    End If

    ' Convert to string
    strValue = CStr(value)

    ' If contains comma, quote, or newline, wrap in quotes and escape quotes
    If InStr(strValue, ",") > 0 Or _
       InStr(strValue, """") > 0 Or _
       InStr(strValue, vbCrLf) > 0 Or _
       InStr(strValue, vbCr) > 0 Or _
       InStr(strValue, vbLf) > 0 Then

        ' Escape quotes by doubling them
        strValue = Replace(strValue, """", """""")

        ' Wrap in quotes
        CSVEscape = """" & strValue & """"
    Else
        CSVEscape = strValue
    End If
End Function

'-------------------------------------------------------------------------------
' FUNCTION: BuildPlantJSON
' PURPOSE: Builds JSON for a single plant's data
' PARAMETERS: strPlantID - Plant identifier
'            totalCables - ByRef counter for total cables
'            totalEndpoints - ByRef counter for total endpoints
' RETURNS: String - JSON representation of plant data
'-------------------------------------------------------------------------------
Private Function BuildPlantJSON(strPlantID As String, _
                               ByRef totalCables As Long, _
                               ByRef totalEndpoints As Long) As String
    On Error GoTo ErrorHandler

    Dim json As String
    Dim ws As Worksheet
    Dim tblCables As ListObject
    Dim tblEndpoints As ListObject
    Dim row As ListRow
    Dim i As Long
    Dim strTableName As String

    ' Start plant object
    json = "    """ & strPlantID & """: {" & vbCrLf

    ' Get appropriate worksheet and tables
    Select Case strPlantID
        Case "WET_PLANT"
            Set ws = sht_WetPlant
            Set tblCables = ws.ListObjects("tbl_WetPlantCables")
            strTableName = "tbl_WetPlantEndpoints"
        Case "ORE_SORTER"
            Set ws = sht_OreSorter
            Set tblCables = ws.ListObjects("tbl_OreSorterCables")
            strTableName = "tbl_OreSorterEndpoints"
        Case "RETREATMENT"
            Set ws = sht_Retreatment
            Set tblCables = ws.ListObjects("tbl_RetreatmentCables")
            strTableName = "tbl_RetreatmentEndpoints"
        Case Else
            BuildPlantJSON = ""
            Exit Function
    End Select

    Set tblEndpoints = sht_Data.ListObjects(strTableName)

    ' Export endpoints
    json = json & "      ""endpoints"": [" & vbCrLf

    If tblEndpoints.ListRows.Count > 0 Then
        For i = 1 To tblEndpoints.ListRows.Count
            Set row = tblEndpoints.ListRows(i)

            json = json & "        {" & vbCrLf
            json = json & "          ""shortName"": """ & JSONEscape(row.Range(1, 1).Value) & """," & vbCrLf
            json = json & "          ""description"": """ & JSONEscape(row.Range(1, 2).Value) & """" & vbCrLf
            json = json & "        }"

            If i < tblEndpoints.ListRows.Count Then
                json = json & ","
            End If
            json = json & vbCrLf

            totalEndpoints = totalEndpoints + 1
        Next i
    End If

    json = json & "      ]," & vbCrLf

    ' Export cables
    json = json & "      ""cables"": [" & vbCrLf

    If tblCables.ListRows.Count > 0 Then
        For i = 1 To tblCables.ListRows.Count
            Set row = tblCables.ListRows(i)

            json = json & "        {" & vbCrLf
            json = json & "          ""scheduled"": " & LCase(CStr(row.Range(1, 1).Value)) & "," & vbCrLf
            json = json & "          ""idAttached"": " & LCase(CStr(row.Range(1, 2).Value)) & "," & vbCrLf
            json = json & "          ""cableID"": """ & JSONEscape(row.Range(1, 3).Value) & """," & vbCrLf
            json = json & "          ""source"": """ & JSONEscape(row.Range(1, 4).Value) & """," & vbCrLf
            json = json & "          ""destination"": """ & JSONEscape(row.Range(1, 5).Value) & """," & vbCrLf
            json = json & "          ""coreSize"": """ & JSONEscape(row.Range(1, 6).Value) & """," & vbCrLf
            json = json & "          ""earthSize"": """ & JSONEscape(row.Range(1, 7).Value) & """," & vbCrLf
            json = json & "          ""coreConfig"": """ & JSONEscape(row.Range(1, 8).Value) & """," & vbCrLf
            json = json & "          ""insulationType"": """ & JSONEscape(row.Range(1, 9).Value) & """," & vbCrLf
            json = json & "          ""cableType"": """ & JSONEscape(row.Range(1, 10).Value) & """," & vbCrLf
            json = json & "          ""cableLength"": """ & JSONEscape(row.Range(1, 11).Value) & """" & vbCrLf
            json = json & "        }"

            If i < tblCables.ListRows.Count Then
                json = json & ","
            End If
            json = json & vbCrLf

            totalCables = totalCables + 1
        Next i
    End If

    json = json & "      ]" & vbCrLf
    json = json & "    }"

    BuildPlantJSON = json
    Exit Function

ErrorHandler:
    BuildPlantJSON = ""
    Debug.Print "Error in BuildPlantJSON: " & Err.Number & " - " & Err.Description
End Function

'-------------------------------------------------------------------------------
' FUNCTION: JSONEscape
' PURPOSE: Escapes special characters for JSON format
' PARAMETERS: value - Value to escape
' RETURNS: String - Escaped value suitable for JSON
'-------------------------------------------------------------------------------
Private Function JSONEscape(value As Variant) As String
    Dim strValue As String

    ' Handle null/empty
    If IsNull(value) Or IsEmpty(value) Then
        JSONEscape = ""
        Exit Function
    End If

    ' Convert to string
    strValue = CStr(value)

    ' Escape special JSON characters
    strValue = Replace(strValue, "\", "\\")     ' Backslash
    strValue = Replace(strValue, """", "\""")   ' Quote
    strValue = Replace(strValue, vbCr, "\r")    ' Carriage return
    strValue = Replace(strValue, vbLf, "\n")    ' Line feed
    strValue = Replace(strValue, vbTab, "\t")   ' Tab

    JSONEscape = strValue
End Function

'-------------------------------------------------------------------------------
' FUNCTION: GetTimestamp
' PURPOSE: Returns formatted timestamp for filenames
' RETURNS: String - Timestamp in format YYYYMMDD_HHNNSS
'-------------------------------------------------------------------------------
Public Function GetTimestamp() As String
    GetTimestamp = Format(Now, "yyyymmdd_hhnnss")
End Function

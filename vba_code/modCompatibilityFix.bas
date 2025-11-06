Attribute VB_Name = "modCompatibilityFix"
' ===============================================================================
' MODULE: modCompatibilityFix
' AUTHOR: AI Assistant
' DATE: 2024-12
' PURPOSE: Automatic compatibility fixing for import operations
'          Handles missing endpoints, column mapping, data conversion
' ===============================================================================

Option Explicit

' Module-level collection to track auto-fixes applied
Private m_autoFixes As Collection

' ===============================================================================
' PUBLIC API - AUTO-FIX FUNCTIONS
' ===============================================================================

'-------------------------------------------------------------------------------
' FUNCTION: InitializeAutoFix
' PURPOSE: Initializes the auto-fix system
'-------------------------------------------------------------------------------
Public Sub InitializeAutoFix()
    Set m_autoFixes = New Collection
End Sub

'-------------------------------------------------------------------------------
' FUNCTION: GetAutoFixReport
' PURPOSE: Returns a report of all auto-fixes applied
' RETURNS: String - Formatted report
'-------------------------------------------------------------------------------
Public Function GetAutoFixReport() As String
    Dim report As String
    Dim fix As Variant
    Dim i As Long

    If m_autoFixes Is Nothing Then
        GetAutoFixReport = "No auto-fixes applied."
        Exit Function
    End If

    If m_autoFixes.Count = 0 Then
        GetAutoFixReport = "No auto-fixes applied."
        Exit Function
    End If

    report = "AUTO-FIXES APPLIED:" & vbCrLf
    report = report & String(50, "=") & vbCrLf

    i = 1
    For Each fix In m_autoFixes
        report = report & i & ". " & fix & vbCrLf
        i = i + 1
    Next fix

    GetAutoFixReport = report
End Function

'-------------------------------------------------------------------------------
' FUNCTION: FixMissingEndpoint
' PURPOSE: Auto-fixes a missing endpoint by creating it with review marker
' PARAMETERS: strPlantID - Plant identifier
'            endpointDesc - Endpoint description that's missing
'            endpointType - "SOURCE" or "DESTINATION"
' RETURNS: String - The short name (created or found)
'-------------------------------------------------------------------------------
Public Function FixMissingEndpoint(strPlantID As String, _
                                  endpointDesc As String, _
                                  endpointType As String) As String
    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim strTableName As String
    Dim newRow As ListRow
    Dim shortName As String
    Dim existingShortName As String

    ' Validate input
    If Trim(endpointDesc) = "" Then
        FixMissingEndpoint = ""
        Exit Function
    End If

    ' First, try to find exact match
    existingShortName = FindEndpointShortName(strPlantID, endpointDesc)
    If existingShortName <> "" Then
        ' Found existing match
        FixMissingEndpoint = existingShortName
        Exit Function
    End If

    ' Try fuzzy match
    existingShortName = FuzzyMatchEndpoint(strPlantID, endpointDesc)
    If existingShortName <> "" Then
        ' Found fuzzy match
        LogAutoFix "Fuzzy matched endpoint: '" & endpointDesc & "' → " & existingShortName
        FixMissingEndpoint = existingShortName
        Exit Function
    End If

    ' No match found - create new endpoint
    ' Get appropriate table
    Select Case strPlantID
        Case "WET_PLANT"
            strTableName = "tbl_WetPlantEndpoints"
            shortName = GenerateShortName(endpointDesc, "1")
        Case "ORE_SORTER"
            strTableName = "tbl_OreSorterEndpoints"
            shortName = GenerateShortName(endpointDesc, "2")
        Case "RETREATMENT"
            strTableName = "tbl_RetreatmentEndpoints"
            shortName = GenerateShortName(endpointDesc, "3")
        Case Else
            FixMissingEndpoint = ""
            Exit Function
    End Select

    Set tbl = sht_Data.ListObjects(strTableName)

    ' Add new endpoint
    Set newRow = tbl.ListRows.Add
    newRow.Range(1, 1).Value = shortName
    newRow.Range(1, 2).Value = endpointDesc & " (Imported - Review)"

    ' Log the auto-fix
    LogAutoFix "Created missing endpoint: " & shortName & " - " & endpointDesc & " (Imported - Review)"

    FixMissingEndpoint = shortName
    Exit Function

ErrorHandler:
    FixMissingEndpoint = ""
    Debug.Print "Error in FixMissingEndpoint: " & Err.Description
End Function

' ===============================================================================
' PRIVATE HELPER FUNCTIONS
' ===============================================================================

'-------------------------------------------------------------------------------
' FUNCTION: FindEndpointShortName
' PURPOSE: Finds an endpoint short name by description (exact match)
' PARAMETERS: strPlantID - Plant identifier
'            description - Description to search for
' RETURNS: String - Short name if found, empty string otherwise
'-------------------------------------------------------------------------------
Private Function FindEndpointShortName(strPlantID As String, description As String) As String
    Dim tbl As ListObject
    Dim strTableName As String
    Dim row As ListRow
    Dim i As Long
    Dim cellDesc As String

    ' Get appropriate table
    Select Case strPlantID
        Case "WET_PLANT": strTableName = "tbl_WetPlantEndpoints"
        Case "ORE_SORTER": strTableName = "tbl_OreSorterEndpoints"
        Case "RETREATMENT": strTableName = "tbl_RetreatmentEndpoints"
        Case Else
            FindEndpointShortName = ""
            Exit Function
    End Select

    Set tbl = sht_Data.ListObjects(strTableName)

    ' Search for exact match (case-insensitive)
    For i = 1 To tbl.ListRows.Count
        Set row = tbl.ListRows(i)
        cellDesc = Trim(CStr(row.Range(1, 2).Value))

        ' Remove "(Imported - Review)" marker if present for comparison
        cellDesc = Replace(cellDesc, " (Imported - Review)", "")

        If StrComp(cellDesc, Trim(description), vbTextCompare) = 0 Then
            FindEndpointShortName = row.Range(1, 1).Value
            Exit Function
        End If
    Next i

    FindEndpointShortName = ""
End Function

'-------------------------------------------------------------------------------
' FUNCTION: FuzzyMatchEndpoint
' PURPOSE: Tries to find close matches for endpoint description
' PARAMETERS: strPlantID - Plant identifier
'            description - Description to search for
' RETURNS: String - Short name of best match, or empty string
'-------------------------------------------------------------------------------
Private Function FuzzyMatchEndpoint(strPlantID As String, description As String) As String
    Dim tbl As ListObject
    Dim strTableName As String
    Dim row As ListRow
    Dim i As Long
    Dim cellDesc As String
    Dim cleanDesc As String
    Dim cleanSearch As String

    ' Get appropriate table
    Select Case strPlantID
        Case "WET_PLANT": strTableName = "tbl_WetPlantEndpoints"
        Case "ORE_SORTER": strTableName = "tbl_OreSorterEndpoints"
        Case "RETREATMENT": strTableName = "tbl_RetreatmentEndpoints"
        Case Else
            FuzzyMatchEndpoint = ""
            Exit Function
    End Select

    Set tbl = sht_Data.ListObjects(strTableName)

    ' Clean search term
    cleanSearch = CleanForFuzzyMatch(description)

    ' Try fuzzy matching
    For i = 1 To tbl.ListRows.Count
        Set row = tbl.ListRows(i)
        cellDesc = Trim(CStr(row.Range(1, 2).Value))
        cellDesc = Replace(cellDesc, " (Imported - Review)", "")
        cleanDesc = CleanForFuzzyMatch(cellDesc)

        ' Check for partial match
        If InStr(1, cleanDesc, cleanSearch, vbTextCompare) > 0 Or _
           InStr(1, cleanSearch, cleanDesc, vbTextCompare) > 0 Then
            FuzzyMatchEndpoint = row.Range(1, 1).Value
            Exit Function
        End If
    Next i

    FuzzyMatchEndpoint = ""
End Function

'-------------------------------------------------------------------------------
' FUNCTION: CleanForFuzzyMatch
' PURPOSE: Cleans a string for fuzzy matching (removes spaces, special chars)
' PARAMETERS: text - Text to clean
' RETURNS: String - Cleaned text
'-------------------------------------------------------------------------------
Private Function CleanForFuzzyMatch(text As String) As String
    Dim result As String

    result = Trim(LCase(text))
    result = Replace(result, " ", "")
    result = Replace(result, "-", "")
    result = Replace(result, "_", "")
    result = Replace(result, ".", "")

    CleanForFuzzyMatch = result
End Function

'-------------------------------------------------------------------------------
' FUNCTION: GenerateShortName
' PURPOSE: Generates a short name from a description
' PARAMETERS: description - Full description
'            plantDigit - Plant digit (1, 2, or 3)
' RETURNS: String - Generated short name (e.g., "XX101")
'-------------------------------------------------------------------------------
Private Function GenerateShortName(description As String, plantDigit As String) As String
    Dim words() As String
    Dim prefix As String
    Dim number As Long
    Dim shortName As String

    ' Try to extract initials from description
    words = Split(Trim(description), " ")

    If UBound(words) >= 1 Then
        ' Multi-word description - use first letters
        prefix = UCase(Left(words(0), 1) & Left(words(1), 1))
    ElseIf Len(words(0)) >= 2 Then
        ' Single word - use first 2 letters
        prefix = UCase(Left(words(0), 2))
    Else
        ' Very short - use "XX"
        prefix = "XX"
    End If

    ' Ensure 2-3 letter prefix
    If Len(prefix) < 2 Then prefix = prefix & "X"
    If Len(prefix) > 3 Then prefix = Left(prefix, 3)

    ' Find next available number for this prefix
    number = FindNextAvailableNumber(prefix, plantDigit)

    ' Build short name
    shortName = prefix & plantDigit & Format(number, "00")

    GenerateShortName = shortName
End Function

'-------------------------------------------------------------------------------
' FUNCTION: FindNextAvailableNumber
' PURPOSE: Finds next available number for a prefix in all endpoint tables
' PARAMETERS: prefix - 2-3 letter prefix
'            plantDigit - Plant digit (1, 2, or 3)
' RETURNS: Long - Next available number (01-19)
'-------------------------------------------------------------------------------
Private Function FindNextAvailableNumber(prefix As String, plantDigit As String) As Long
    Dim strPlantID As String
    Dim tbl As ListObject
    Dim strTableName As String
    Dim row As ListRow
    Dim i As Long
    Dim existingName As String
    Dim pattern As String
    Dim number As Long
    Dim maxNum As Long

    ' Determine plant
    Select Case plantDigit
        Case "1": strPlantID = "WET_PLANT": strTableName = "tbl_WetPlantEndpoints"
        Case "2": strPlantID = "ORE_SORTER": strTableName = "tbl_OreSorterEndpoints"
        Case "3": strPlantID = "RETREATMENT": strTableName = "tbl_RetreatmentEndpoints"
        Case Else
            FindNextAvailableNumber = 1
            Exit Function
    End Select

    Set tbl = sht_Data.ListObjects(strTableName)

    ' Find highest number for this prefix
    maxNum = 0
    pattern = UCase(prefix) & plantDigit

    For i = 1 To tbl.ListRows.Count
        Set row = tbl.ListRows(i)
        existingName = UCase(Trim(CStr(row.Range(1, 1).Value)))

        ' Check if matches pattern
        If Left(existingName, Len(pattern)) = pattern Then
            ' Extract number
            number = Val(Mid(existingName, Len(pattern) + 1, 2))
            If number > maxNum Then maxNum = number
        End If
    Next i

    ' Return next number (max 19 to stay within 00-19 range)
    FindNextAvailableNumber = maxNum + 1
    If FindNextAvailableNumber > 19 Then FindNextAvailableNumber = 19
End Function

'-------------------------------------------------------------------------------
' SUB: LogAutoFix
' PURPOSE: Logs an auto-fix action
' PARAMETERS: description - Description of the fix
'-------------------------------------------------------------------------------
Private Sub LogAutoFix(description As String)
    If m_autoFixes Is Nothing Then
        Set m_autoFixes = New Collection
    End If

    m_autoFixes.Add description
End Sub

'-------------------------------------------------------------------------------
' FUNCTION: NormalizePlantID
' PURPOSE: Normalizes plant identifier variations
' PARAMETERS: plantID - Plant identifier (any variation)
' RETURNS: String - Normalized plant ID
'-------------------------------------------------------------------------------
Public Function NormalizePlantID(plantID As String) As String
    Dim normalized As String

    normalized = UCase(Trim(plantID))
    normalized = Replace(normalized, " ", "_")
    normalized = Replace(normalized, "-", "_")

    Select Case normalized
        Case "WETPLANT", "WET_PLANT", "WET", "PLANT1", "1"
            NormalizePlantID = "WET_PLANT"
        Case "ORESORTER", "ORE_SORTER", "ORE", "PLANT2", "2"
            NormalizePlantID = "ORE_SORTER"
        Case "RETREATMENT", "RETREATMENT_PLANT", "RETREAT", "PLANT3", "3"
            NormalizePlantID = "RETREATMENT"
        Case Else
            NormalizePlantID = normalized
    End Select
End Function

'-------------------------------------------------------------------------------
' FUNCTION: ConvertDataType
' PURPOSE: Converts data to appropriate type with error handling
' PARAMETERS: value - Value to convert
'            targetType - "BOOLEAN", "NUMBER", "TEXT"
' RETURNS: Variant - Converted value
'-------------------------------------------------------------------------------
Public Function ConvertDataType(value As Variant, targetType As String) As Variant
    On Error Resume Next

    Select Case UCase(targetType)
        Case "BOOLEAN"
            ' Handle various boolean representations
            If IsNumeric(value) Then
                ConvertDataType = CBool(value)
            ElseIf VarType(value) = vbBoolean Then
                ConvertDataType = value
            Else
                Dim strVal As String
                strVal = UCase(Trim(CStr(value)))
                Select Case strVal
                    Case "TRUE", "YES", "Y", "1", "ON"
                        ConvertDataType = True
                    Case "FALSE", "NO", "N", "0", "OFF", ""
                        ConvertDataType = False
                    Case Else
                        ConvertDataType = False
                End Select
            End If

        Case "NUMBER"
            ' Strip non-numeric characters and convert
            Dim numStr As String
            Dim char As String
            Dim i As Long
            numStr = CStr(value)

            ' Remove common suffixes (m, mm, etc.)
            numStr = Replace(numStr, "m", "")
            numStr = Replace(numStr, "M", "")

            ConvertDataType = Val(numStr)

        Case "TEXT"
            ConvertDataType = CStr(value)

        Case Else
            ConvertDataType = value
    End Select

    If Err.Number <> 0 Then
        ConvertDataType = value
        Err.Clear
    End If
End Function

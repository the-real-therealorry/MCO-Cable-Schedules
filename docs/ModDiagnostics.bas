Attribute VB_Name = "ModDiagnostics"
' ===============================================================================
' MODULE: ModDiagnostics
' PURPOSE: Diagnostic tools to help debug the cable edit endpoint lookup failure
' ===============================================================================

Option Explicit

' ===============================================================================
' SUBROUTINE: DiagnoseEndpointLookup
' PURPOSE: Diagnoses why endpoint lookup is failing for a specific cable
' USAGE: Call this from Immediate Window:  DiagnoseEndpointLookup "CV103-C1001-CV103", "WET_PLANT"
' ===============================================================================
Public Sub DiagnoseEndpointLookup(cableID As String, strFormID As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "==============================================="
    Debug.Print "ENDPOINT LOOKUP DIAGNOSTIC"
    Debug.Print "==============================================="
    Debug.Print "Cable ID: " & cableID
    Debug.Print "Plant Type: " & strFormID
    Debug.Print ""
    
    ' Get the cable object
    Dim tableIndex As Long
    tableIndex = FindCableRowIndex(cableID, strFormID)
    
    If tableIndex = 0 Then
        Debug.Print "? ERROR: Cable not found in table"
        Exit Sub
    End If
    
    Debug.Print "? Cable found at table row index: " & tableIndex
    Debug.Print ""
    
    ' Get cable data
    Dim cCable As clCable
    Set cCable = modDatabase.GetCableByRowIndex(strFormID, tableIndex)
    
    If cCable Is Nothing Then
        Debug.Print "? ERROR: Could not load cable data"
        Exit Sub
    End If
    
    Debug.Print "CABLE DATA FROM TABLE:"
    Debug.Print "  Source (stored): [" & cCable.Source & "]"
    Debug.Print "  Destination (stored): [" & cCable.Destination & "]"
    Debug.Print ""
    
    ' Get all endpoints for this plant
    Debug.Print "LOADING ENDPOINTS FOR " & strFormID & "..."
    Dim arrEndpoints As Variant
    arrEndpoints = sht_Data.GetEndpointsArray(strFormID)
    
    If Not IsArray(arrEndpoints) Then
        Debug.Print "? ERROR: Could not load endpoints array"
        Exit Sub
    End If
    
    Dim endpointCount As Long
    On Error Resume Next
    endpointCount = UBound(arrEndpoints) - LBound(arrEndpoints) + 1
    On Error GoTo ErrorHandler
    
    Debug.Print "? Found " & endpointCount & " endpoints"
    Debug.Print ""
    
    ' List all endpoints
    Debug.Print "ALL ENDPOINTS IN TABLE:"
    Debug.Print "Index | Short Name | Description"
    Debug.Print "------|------------|-------------"
    
    Dim i As Long
    Dim epItem As Variant
    For Each epItem In arrEndpoints
        Debug.Print "  " & i & "   | " & epItem.ShortName & " | " & epItem.description
        i = i + 1
    Next epItem
    Debug.Print ""
    
    ' Try to find source endpoint
    Debug.Print "SEARCHING FOR SOURCE ENDPOINT..."
    Debug.Print "  Looking for description: [" & cCable.Source & "]"
    
    Dim foundSource As Boolean
    foundSource = False
    Dim sourceShortName As String
    
    For Each epItem In arrEndpoints
        If StrComp(epItem.description, Trim(cCable.Source), vbTextCompare) = 0 Then
            foundSource = True
            sourceShortName = epItem.ShortName
            Debug.Print "  ? FOUND: Short Name = " & sourceShortName
            Exit For
        End If
    Next epItem
    
    If Not foundSource Then
        Debug.Print "  ? NOT FOUND"
        Debug.Print "  This is why the Source dropdown appears empty!"
        Debug.Print ""
        Debug.Print "  CHECKING FOR SIMILAR MATCHES..."
        For Each epItem In arrEndpoints
            If InStr(1, epItem.description, Trim(cCable.Source), vbTextCompare) > 0 Or _
               InStr(1, Trim(cCable.Source), epItem.description, vbTextCompare) > 0 Then
                Debug.Print "    Similar: " & epItem.ShortName & " - " & epItem.description
            End If
        Next epItem
    End If
    Debug.Print ""
    
    ' Try to find destination endpoint
    Debug.Print "SEARCHING FOR DESTINATION ENDPOINT..."
    Debug.Print "  Looking for description: [" & cCable.Destination & "]"
    
    Dim foundDest As Boolean
    foundDest = False
    Dim destShortName As String
    
    For Each epItem In arrEndpoints
        If StrComp(epItem.description, Trim(cCable.Destination), vbTextCompare) = 0 Then
            foundDest = True
            destShortName = epItem.ShortName
            Debug.Print "  ? FOUND: Short Name = " & destShortName
            Exit For
        End If
    Next epItem
    
    If Not foundDest Then
        Debug.Print "  ? NOT FOUND"
        Debug.Print "  This is why the Destination dropdown appears empty!"
        Debug.Print ""
        Debug.Print "  CHECKING FOR SIMILAR MATCHES..."
        For Each epItem In arrEndpoints
            If InStr(1, epItem.description, Trim(cCable.Destination), vbTextCompare) > 0 Or _
               InStr(1, Trim(cCable.Destination), epItem.description, vbTextCompare) > 0 Then
                Debug.Print "    Similar: " & epItem.ShortName & " - " & epItem.description
            End If
        Next epItem
    End If
    Debug.Print ""
    
    ' Provide recommendations
    Debug.Print "DIAGNOSIS SUMMARY:"
    Debug.Print "==============================================="
    
    If foundSource And foundDest Then
        Debug.Print "? Both endpoints found - lookup should work"
        Debug.Print "  If edit form still shows empty, check GetShortNameFromDescription function"
    Else
        Debug.Print "? Problem identified:"
        If Not foundSource Then
            Debug.Print "  - Source endpoint description doesn't match any endpoint in table"
        End If
        If Not foundDest Then
            Debug.Print "  - Destination endpoint description doesn't match any endpoint in table"
        End If
        Debug.Print ""
        Debug.Print "SOLUTION OPTIONS:"
        Debug.Print "1. Add the missing endpoint(s) to the endpoints table"
        Debug.Print "2. Update the cable record to use an existing endpoint"
        Debug.Print "3. Fix any spelling/spacing differences in descriptions"
    End If
    
    Debug.Print "==============================================="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "? ERROR in diagnostic: " & Err.description
End Sub

' ===============================================================================
' FUNCTION: FindCableRowIndex
' PURPOSE: Finds the table row index for a cable by its ID
' ===============================================================================
Private Function FindCableRowIndex(cableID As String, strFormID As String) As Long
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    
    ' Determine which worksheet and table to search
    Select Case strFormID
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
            FindCableRowIndex = 0
            Exit Function
    End Select
    
    ' Search for cable ID in column 3 (Cable ID column)
    For i = 1 To tbl.ListRows.count
        If StrComp(tbl.ListRows(i).Range(1, 3).Value, cableID, vbTextCompare) = 0 Then
            FindCableRowIndex = i
            Exit Function
        End If
    Next i
    
    ' Not found
    FindCableRowIndex = 0
    
    Exit Function
    
ErrorHandler:
    FindCableRowIndex = 0
End Function

' ===============================================================================
' SUBROUTINE: ListAllEndpoints
' PURPOSE: Lists all endpoints for a plant type
' USAGE: Call from Immediate Window:  ListAllEndpoints "WET_PLANT"
' ===============================================================================
Public Sub ListAllEndpoints(strFormID As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "==============================================="
    Debug.Print "ALL ENDPOINTS FOR " & strFormID
    Debug.Print "==============================================="
    
    Dim arrEndpoints As Variant
    arrEndpoints = sht_Data.GetEndpointsArray(strFormID)
    
    If Not IsArray(arrEndpoints) Then
        Debug.Print "? ERROR: Could not load endpoints"
        Exit Sub
    End If
    
    Debug.Print "Short Name | Description"
    Debug.Print "-----------|-------------"
    
    Dim epItem As Variant
    For Each epItem In arrEndpoints
        Debug.Print epItem.ShortName & " | " & epItem.description
    Next epItem
    
    Debug.Print "==============================================="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "? ERROR: " & Err.description
End Sub

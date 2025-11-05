Attribute VB_Name = "sht_WetPlant"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' ==============================================================================
' MODULE: sht_WetPlant (Worksheet Class Module)
' AUTHOR: jorr@mtcarbine.com.au
' DATE: 20-12-24
' PURPOSE: Manages cable data for the Wet Screen Crushing Plant. Provides CRUD
'          operations for cable records stored in the tbl_WetPlantCables Excel
'          table. Handles cable scheduling, circuit numbering, and data conversion
'          for the wet processing facility.
' ==============================================================================

Option Explicit

' ==============================================================================
' MODULE-LEVEL VARIABLES
' ==============================================================================

' Tracks the first empty row number (currently unused but may be for future functionality)
Dim lngFirstEmptyRow As Long

' ==============================================================================
' PUBLIC FUNCTIONS - TABLE ROW MANAGEMENT
' ==============================================================================

'------------------------------------------------------------------------------
' FUNCTION: GetNextEmptyRowNumber
' PURPOSE: Returns the row number where the next cable record should be added
' RETURNS: Long - Row number (table row count + 2 to account for header)
' NOTES: Adds 2 because ListRows.count doesn't include header, but we need
'        actual worksheet row number
'------------------------------------------------------------------------------
Public Function GetNextEmptyRowNumber() As Long
    On Error GoTo ErrorHandler

    ' Add 2: +1 for header row, +1 for next empty row
    GetNextEmptyRowNumber = Me.ListObjects("tbl_WetPlantCables").ListRows.count + 2

ErrorExit:
Exit Function

ErrorHandler:
    ' Standardized error handling using HandleError procedure
    HandleError "sht_WetPlant", "GetNextEmptyRowNumber", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function

'------------------------------------------------------------------------------
' FUNCTION: GetLastRowNumber
' PURPOSE: Returns the row number of the last data row in the table
' RETURNS: Long - Last row number with data (table row count + 1 for header)
' NOTES: Different from GetNextEmptyRowNumber by 1 row
'------------------------------------------------------------------------------
Public Function GetLastRowNumber() As Long
    On Error GoTo ErrorHandler

    ' Add 1 to account for header row
    GetLastRowNumber = Me.ListObjects("tbl_WetPlantCables").ListRows.count + 1

ErrorExit:
Exit Function

ErrorHandler:
    ' Standardized error handling using HandleError procedure
    HandleError "sht_WetPlant", "GetLastRowNumber", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function

'------------------------------------------------------------------------------
' FUNCTION: GetNumberOfCables
' PURPOSE: Returns the total number of cable records in the table
' RETURNS: Long - Count of cable records (excluding header)
' NOTES: Simple wrapper around ListRows.count for consistency
'------------------------------------------------------------------------------
Public Function GetNumberOfCables() As Long
    On Error GoTo ErrorHandler

    GetNumberOfCables = Me.ListObjects("tbl_WetPlantCables").ListRows.count

ErrorExit:
Exit Function

ErrorHandler:
    ' Standardized error handling using HandleError procedure
    HandleError "sht_WetPlant", "GetNumberOfCables", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function

' ==============================================================================
' PUBLIC FUNCTIONS - CABLE DATA MANAGEMENT
' ==============================================================================

'------------------------------------------------------------------------------
' FUNCTION: GetCablesArray
' PURPOSE: Retrieves all cable records as an array of clCable objects
' RETURNS: Variant - Array of clCable objects containing all cable data
' NOTES: Handles empty table case and maps table columns to object properties
'        Uses column constants (cc*) that should be defined elsewhere
'        Uses .Scheduled property (consistent with Retreatment module)
'------------------------------------------------------------------------------
Public Function GetCablesArray() As Variant
    On Error GoTo ErrorHandler
    
    Dim cTemp As New clCable
    Dim arrRows As Variant
    Dim arrCables() As New clCable
    
    ' Handle empty table case - return single empty element array
    If Me.GetNumberOfCables = 0 Then
        ReDim arrCables(1 To 1)
        GetCablesArray = arrCables
        Exit Function
        
    Else
        ' Get raw data from table (excludes header automatically)
        arrRows = sht_WetPlant.ListObjects("tbl_WetPlantCables").DataBodyRange.Value
        ReDim arrCables(1 To UBound(arrRows, 1))
        
    End If
    
    Dim lngRow As Long, lngCol As Long
    
    ' Process each row of data (header already excluded by DataBodyRange)
    For lngRow = 1 To UBound(arrRows, 1)
        ' Process each column and map to appropriate cable object property
        For lngCol = 1 To UBound(arrRows, 2)
            ' Map columns to cable object properties using column constants
            If lngCol = ccScheduled Then
                arrCables(lngRow).Scheduled = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = ccIDAttached Then
                arrCables(lngRow).IDAttached = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = ccCableID Then
                arrCables(lngRow).cableID = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = ccSource Then
                arrCables(lngRow).Source = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = ccDestination Then
                arrCables(lngRow).Destination = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = ccCoreSize Then
                arrCables(lngRow).CoreSize = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = ccEarthsize Then
                arrCables(lngRow).EarthSize = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = cccoreconfig Then
                arrCables(lngRow).CoreConfig = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = ccinsulationtype Then
                arrCables(lngRow).InsulationType = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = cccabletype Then
                arrCables(lngRow).CableType = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = cccablelength Then
                arrCables(lngRow).CableLength = arrRows(lngRow, lngCol)
                
            End If
        
        Next lngCol
    Next lngRow
        
    ' Return the populated array of cable objects
    GetCablesArray = arrCables
    
    ' Clean up object references
    Set arrRows = Nothing
    
ErrorExit:
Exit Function

ErrorHandler:
    ' Standardized error handling using HandleError procedure
    Set arrRows = Nothing
    HandleError "sht_WetPlant", "GetCablesArray", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function

'------------------------------------------------------------------------------
' FUNCTION: ConvertCablesToRows
' PURPOSE: Converts an array of clCable objects back to a 2D array for table operations
' PARAMETERS: arrCables - Array of clCable objects to convert
' RETURNS: Variant - 2D array with cable data in table format
' NOTES: Inverse operation of GetCablesArray. Correctly uses .Scheduled property
'        (consistent with GetCablesArray, unlike OreSorter module)
'------------------------------------------------------------------------------
Public Function ConvertCablesToRows(ByVal arrCables As Variant) As Variant
    On Error GoTo ErrorHandler
    
    ' Create 2D array sized for cables count by total columns
    Dim arrRows() As Variant
    ReDim arrRows(1 To UBound(arrCables, 1), 1 To ccTotalColumns)
    
    Dim lngRow As Long
    Dim lngCol As Long
    
    ' Convert each cable object back to row format
    For lngRow = 1 To UBound(arrRows, 1)
        For lngCol = 1 To UBound(arrRows, 2)
            ' Map cable object properties back to column positions
            ' NOTE: Correctly uses .Scheduled (consistent with GetCablesArray)
            If lngCol = ccScheduled Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).Scheduled
                
            ElseIf lngCol = ccIDAttached Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).IDAttached
                
            ElseIf lngCol = ccCableID Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).cableID
                
            ElseIf lngCol = ccSource Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).Source
                
            ElseIf lngCol = ccDestination Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).Destination
                
            ElseIf lngCol = ccCoreSize Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).CoreSize
                
            ElseIf lngCol = ccEarthsize Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).EarthSize
                
            ElseIf lngCol = cccoreconfig Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).CoreConfig
                
            ElseIf lngCol = ccinsulationtype Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).InsulationType
                
            ElseIf lngCol = cccabletype Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).CableType
                
            ElseIf lngCol = cccablelength Then
                arrRows(lngRow, lngCol) = arrCables(lngRow).CableLength
                
            End If
        Next lngCol
    Next lngRow
                
    ' Return the converted 2D array
    ConvertCablesToRows = arrRows
    
    ' Clean up object references
    Set arrCables = Nothing
    
ErrorExit:
Exit Function

ErrorHandler:
    ' Standardized error handling using HandleError procedure
    Set arrCables = Nothing
    HandleError "sht_WetPlant", "ConvertCablesToRows", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function

'------------------------------------------------------------------------------
' FUNCTION: SaveCable
' PURPOSE: Saves a new cable record to the table
' PARAMETERS: cNewCable - clCable object containing the cable data to save
' RETURNS: Long - The total number of cables after adding the new one
' NOTES: Uses the cable object's ToRow() method to get array representation
'        Hard-coded to 11 columns - should use constant for maintainability
'------------------------------------------------------------------------------
Public Function SaveCable(ByVal cNewCable As clCable) As Long
    On Error GoTo ErrorHandler
    
    ' Get cable data as array using the object's ToRow method
    Dim arrRow() As Variant
    arrRow = cNewCable.ToRow()
    
    ' Add new row to the table
    Dim lrNewRow As ListRow
    Set lrNewRow = Me.ListObjects("tbl_WetPlantCables").ListRows.Add
    
    ' Populate all columns with cable data
    Dim i As Long
    
    For i = 1 To ccTotalColumns
        lrNewRow.Range(i) = arrRow(i)
    Next i
        
    ' Return the total number of cables (including the new one)
    Dim lngNewRowNumber As Long
    lngNewRowNumber = Me.GetNumberOfCables()
    
    SaveCable = lngNewRowNumber

    ' Apply checkbox formatting to Boolean columns
    Call ApplyCheckboxFormatting

ErrorExit:
Exit Function

ErrorHandler:
    ' Standardized error handling using HandleError procedure
    HandleError "sht_WetPlant", "SaveCable", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function

' ==============================================================================
' PUBLIC FUNCTIONS - CIRCUIT NUMBER MANAGEMENT
' ==============================================================================

'------------------------------------------------------------------------------
' FUNCTION: GetNextCircuitNumber
' PURPOSE: Calculates the next available circuit number by finding the highest existing one
' RETURNS: Long - Next sequential circuit number to use
' NOTES: Uses regex to extract circuit numbers from cable IDs and finds the maximum
'        Starting number is 1001, indicating Wet Plant designation (first plant type)
'        Processes column C which should contain CableID values
'        Pattern uses "1" to identify Wet Plant endpoints
'------------------------------------------------------------------------------
Public Function GetNextCircuitNumber() As Long
    On Error GoTo ErrorHandler

    Dim lngCurrent As Long
    Dim lngLastRow As Long
    Dim lngNextCircuitNumber As Long
    
    Dim lngHighest As Long
    lngHighest = 1001  ' (2001 for OreSorter, 3001 for Retreatment)
    
    Dim strPattern As String
    strPattern = "[A-Z]{2,3}[1][0-1][0-9]"
    
    Dim strItem As String
    Dim strCurrent As String
    
    lngLastRow = Me.GetLastRowNumber()
    
    Dim rItem As Range
    If lngLastRow > 1 Then
        For Each rItem In sht_WetPlant.Range("C2:C" & lngLastRow)
        
            ' DEFENSIVE: Handle empty, null, or error values
            If Not IsEmpty(rItem.Value) And Not IsNull(rItem.Value) And Not IsError(rItem.Value) Then
                strItem = Trim(CStr(rItem.Value))
                
                ' Only process if we have actual content
                If Len(strItem) > 0 Then
                    strCurrent = modUtils.RegexReplace(strItem, strPattern, "")
                    strCurrent = modUtils.RemoveChar(strCurrent, "-")
                    strCurrent = modUtils.RemoveChar(strCurrent, "C")
                    
                    ' Ensure result is numeric before converting
                    If IsNumeric(strCurrent) And Len(strCurrent) > 0 Then
                        lngCurrent = CLng(strCurrent)
                        
                        If lngCurrent > lngHighest Then
                            lngHighest = lngCurrent
                        End If
                    End If
                End If
            End If
        Next rItem
    
        lngNextCircuitNumber = lngHighest + 1
        
    Else
        lngNextCircuitNumber = lngHighest
        
    End If
    
    GetNextCircuitNumber = lngNextCircuitNumber
    
ErrorExit:
Exit Function

ErrorHandler:
    HandleError "sht_WetPlant", "GetNextCircuitNumber", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function

' ==============================================================================
' PUBLIC SUBROUTINES - TABLE FORMATTING
' ==============================================================================

'------------------------------------------------------------------------------
' SUBROUTINE: ResetTableStyle
' PURPOSE: Resets the table formatting by clearing and reapplying the table style
' NOTES: Uses TableStyleMedium16 (different from other plants' styles)
'        Added error handling for consistency with other modules
'------------------------------------------------------------------------------
Public Sub ResetTableStyle()
On Error GoTo AIError

    ' Clear existing style then reapply to ensure clean formatting
    Me.ListObjects("tbl_WetPlantCables").TableStyle = ""
    Me.ListObjects("tbl_WetPlantCables").TableStyle = "TableStyleMedium16"

AIExit:
Exit Sub

AIError:
    ' Uses HandleError procedure for consistent error handling
    HandleError "sht_WetPlant", "ResetTableStyle", Err.Number, Err.description, Erl
    Resume AIExit
    
End Sub

' ==============================================================================
' PRIVATE EVENT HANDLERS - WORKSHEET EVENTS
' ==============================================================================

'------------------------------------------------------------------------------
' EVENT HANDLER: Worksheet_SelectionChange
' PURPOSE: Handles user selection changes to manage row action buttons
' PARAMETERS: Target - The range that was selected
' NOTES: Only processes when worksheet is visible, delegates to ModRowActions
'        for button management logic. Already uses HandleError procedure.
'------------------------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo CleanFail
    
    ' Only process selection changes when worksheet is visible to user
    If Me.Visible <> xlSheetVisible Then Exit Sub
    
    ' Delegate to external module for row button management
    ModRowActions.OnSelectionChangedForRowButtons Me, Target
    
    Exit Sub
    
CleanFail:
    ' Error handling with descriptive module name for context
    HandleError "Wet Screen Crushing Plant", "Worksheet_SelectionChange", Err.Number, Err.description, Erl
    
End Sub

'------------------------------------------------------------------------------
' EVENT HANDLER: Worksheet_Deactivate
' PURPOSE: Cleans up row action buttons when user leaves this worksheet
' NOTES: Uses On Error Resume Next since this is cleanup code that shouldn't
'        interrupt user workflow if it fails
'------------------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    
    ' Hide row action buttons when leaving the worksheet
    ModRowActions.HideRowButtons Me
    
End Sub

' ==============================================================================
' ADD THESE METHODS TO EACH PLANT WORKSHEET MODULE
' (sht_WetPlant, sht_OreSorter, sht_Retreatment)
' ==============================================================================

' ==============================================================================
' FOR sht_WetPlant.cls
' ==============================================================================

'------------------------------------------------------------------------------
' SUBROUTINE: DeleteCableRow
' PURPOSE: Properly deletes a cable row from the Wet Plant cables table
' PARAMETERS: lngRowIndex - The row index in the ListRows collection (1-based)
' NOTES: Use this instead of manually deleting rows to maintain table integrity
'        This ensures the table structure remains valid after deletion
'------------------------------------------------------------------------------
Public Sub DeleteCableRow(lngRowIndex As Long)
    On Error GoTo ErrorHandler
    
    ' Validate row index exists within table bounds
    If lngRowIndex < 1 Or lngRowIndex > Me.ListObjects("tbl_WetPlantCables").ListRows.count Then
        MsgBox "Invalid row index: " & lngRowIndex, vbExclamation
        Exit Sub
    End If
    
    ' Delete the entire ListRow properly (not just clear contents)
    Me.ListObjects("tbl_WetPlantCables").ListRows(lngRowIndex).Delete
    
    ' Optional: Show confirmation
    ' MsgBox "Cable deleted successfully.", vbInformation
    
ErrorExit:
Exit Sub

ErrorHandler:
    HandleError "sht_WetPlant", "DeleteCableRow", Err.Number, Err.description, Erl
    Resume ErrorExit
End Sub

'------------------------------------------------------------------------------
' FUNCTION: GetCableByRowIndex
' PURPOSE: Retrieves a cable object from a specific table row
' PARAMETERS: lngRowIndex - The row index in the ListRows collection (1-based)
' RETURNS: clCable object populated with data from the specified row, or Nothing if invalid
' NOTES: This method allows retrieval of a single cable for editing purposes
'------------------------------------------------------------------------------
Public Function GetCableByRowIndex(lngRowIndex As Long) As clCable
    On Error GoTo ErrorHandler
    
    ' Validate row index exists within table bounds
    If lngRowIndex < 1 Or lngRowIndex > Me.ListObjects("tbl_WetPlantCables").ListRows.count Then
        MsgBox "Invalid row index: " & lngRowIndex, vbExclamation
        Set GetCableByRowIndex = Nothing
        Exit Function
    End If
    
    ' Create new cable object to populate
    Dim cCable As New clCable
    
    ' Get the data from the specified row
    Dim arrData As Variant
    arrData = Me.ListObjects("tbl_WetPlantCables").ListRows(lngRowIndex).Range.Value
    
    ' Map table columns to cable object properties
    With cCable
        .Scheduled = arrData(1, ccScheduled)
        .IDAttached = arrData(1, ccIDAttached)
        .cableID = arrData(1, ccCableID)
        .Source = arrData(1, ccSource)
        .Destination = arrData(1, ccDestination)
        .CoreSize = arrData(1, ccCoreSize)
        .EarthSize = arrData(1, ccEarthsize)
        .CoreConfig = arrData(1, cccoreconfig)
        .InsulationType = arrData(1, ccinsulationtype)
        .CableType = arrData(1, cccabletype)
        .CableLength = arrData(1, cccablelength)
    End With
    
    ' Return the populated cable object
    Set GetCableByRowIndex = cCable
    
ErrorExit:
Exit Function

ErrorHandler:
    HandleError "sht_WetPlant", "GetCableByRowIndex", Err.Number, Err.description, Erl
    Set GetCableByRowIndex = Nothing
    Resume ErrorExit
End Function

'------------------------------------------------------------------------------
' FUNCTION: UpdateCable
' PURPOSE: Updates an existing cable record in the table
' PARAMETERS: lngRowIndex - The row index in the ListRows collection (1-based)
'            cUpdatedCable - clCable object containing the updated cable data
' RETURNS: Boolean - True if update successful, False if failed
' NOTES: Replaces entire row with updated data from cable object
'------------------------------------------------------------------------------
Public Function UpdateCable(lngRowIndex As Long, cUpdatedCable As clCable) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate row index exists within table bounds
    If lngRowIndex < 1 Or lngRowIndex > Me.ListObjects("tbl_WetPlantCables").ListRows.count Then
        MsgBox "Invalid row index: " & lngRowIndex, vbExclamation
        UpdateCable = False
        Exit Function
    End If
    
    ' Update the row with new cable data
    Dim arrRow As Variant
    arrRow = cUpdatedCable.ToRow()
    
    ' Write the updated data to the table row
    Me.ListObjects("tbl_WetPlantCables").ListRows(lngRowIndex).Range.Value = arrRow
    
    ' Return success
    UpdateCable = True

    ' Apply checkbox formatting to Boolean columns
    Call ApplyCheckboxFormatting

ErrorExit:
Exit Function

ErrorHandler:
    HandleError "sht_WetPlant", "UpdateCable", Err.Number, Err.description, Erl
    UpdateCable = False
    Resume ErrorExit
End Function

' ==============================================================================
' CHECKBOX FORMATTING FIX
' Purpose: Applies Excel's checkbox data type to Boolean columns
' Author: Added for checkbox display fix
' Date: 2025-11-05
' ==============================================================================

'------------------------------------------------------------------------------
' SUBROUTINE: ApplyCheckboxFormatting
' PURPOSE: Converts Boolean TRUE/FALSE values in columns 1-2 to display as checkboxes
' NOTES: Uses Excel's built-in checkbox data type feature (Excel 365/2019+)
'        Falls back to refresh method for older Excel versions
'------------------------------------------------------------------------------
Private Sub ApplyCheckboxFormatting()
    On Error Resume Next

    Dim tblCables As ListObject
    Dim rngScheduled As Range
    Dim rngIDAttached As Range

    Set tblCables = Me.ListObjects("tbl_WetPlantCables")

    ' Only apply if table has data
    If Not tblCables Is Nothing Then
        If tblCables.ListRows.count > 0 Then
            ' Get ranges for the two checkbox columns
            Set rngScheduled = tblCables.ListColumns(1).DataBodyRange
            Set rngIDAttached = tblCables.ListColumns(2).DataBodyRange

            ' Method 1: Try Excel 365 Checkbox Data Type
            On Error Resume Next
            rngScheduled.ExcelDataType = xlCheckbox
            rngIDAttached.ExcelDataType = xlCheckbox

            ' Method 2: If Method 1 fails (xlCheckbox constant not available), refresh cells
            If Err.Number <> 0 Then
                Err.Clear
                ' Force Excel to re-evaluate the cell format
                Dim cell As Range
                For Each cell In rngScheduled
                    If cell.Value = True Or cell.Value = False Then
                        cell.Value = cell.Value  ' Refresh display
                    End If
                Next cell

                For Each cell In rngIDAttached
                    If cell.Value = True Or cell.Value = False Then
                        cell.Value = cell.Value  ' Refresh display
                    End If
                Next cell
            End If
            On Error GoTo 0
        End If
    End If

    ' Cleanup
    Set rngScheduled = Nothing
    Set rngIDAttached = Nothing
    Set tblCables = Nothing
End Sub

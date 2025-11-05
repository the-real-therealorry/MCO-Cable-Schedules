# VBA Code Analysis: Checkbox Data Handling in Power Cable Register

## Summary
The checkbox data for "Cable Scheduled" and "Cable Label Attached" columns is handled through a combination of:
1. Boolean properties in the `clCable` class (storage)
2. Checkbox controls in the `frm_RegisterCable` form (UI)
3. Update methods in worksheet modules and database layer

---

## 1. DATA MODEL - clCable Class

### File Location: `xl/vbaProject.bin` (OLE Stream: `VBA/clCable`)

The `clCable` class defines two Boolean properties for checkbox columns:

```vba
' Private member variables
Private mIsScheduled As Boolean
Private mIDAttached As Boolean

' Column enumeration
Public Enum eCableColumns
    ccScheduled = 1          ' Column 1: Scheduled checkbox
    ccIDAttached = 2         ' Column 2: ID Attached checkbox
    ccCableID = 3
    ccSource = 4
    ccDestination = 5
    ccCoreSize = 6
    ccEarthsize = 7
    cccoreconfig = 8
    ccinsulationtype = 9
    cccabletype = 10
    cccablelength = 11
    ccTotalColumns = 11
End Enum

' Property: Scheduled (Boolean checkbox)
Public Property Get Scheduled() As Boolean
    Scheduled = mIsScheduled
End Property

Public Property Let Scheduled(bIsScheduled As Boolean)
    mIsScheduled = bIsScheduled
End Property

' Property: IDAttached (Boolean checkbox)
Public Property Get IDAttached() As Boolean
    IDAttached = mIDAttached
End Property

Public Property Let IDAttached(bIDAttached As Boolean)
    mIDAttached = bIDAttached
End Property
```

### ToRow() Method - Converts Object to Array

```vba
Public Function ToRow() As Variant
    On Error GoTo ErrorHandler
    
    Dim arrRow As Variant
    ReDim arrRow(1 To 11)
    
    Dim lngRow As Long
    For lngRow = 1 To ccTotalColumns
        If lngRow = ccScheduled Then
            arrRow(lngRow) = mIsScheduled        ' Write Boolean value
            
        ElseIf lngRow = ccIDAttached Then
            arrRow(lngRow) = mIDAttached         ' Write Boolean value
        
        ElseIf lngRow = ccCableID Then
            arrRow(lngRow) = mCableID
                    
        ' ... other columns ...
        
        End If
    Next lngRow
    
    ToRow = arrRow
    
done:
Exit Function
ErrorHandler:
    MsgBox "An error occurred: " & Err.description, vbCritical
    Resume Next
End Function
```

---

## 2. UI FORM - frm_RegisterCable

### File Location: `xl/vbaProject.bin` (OLE Stream: `VBA/frm_RegisterCable`)

#### Checkbox Controls
The form has two checkbox controls:
- `cb_IsScheduled` - "Cable Scheduled Complete"
- `cb_IDAttached` - "Cable Label Attached"

#### ReloadForm() - Resets Checkboxes

```vba
Public Sub ReloadForm()
    On Error GoTo ErrorHandler

    ' ... other field resets ...

    ' RESET ALL CHECKBOXES TO UNCHECKED
    Me.cb_IsScheduled.Value = False
    Me.cb_IDAttached.Value = False

    ' ... rest of method ...
    
ErrorHandler_Exit:
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred resetting the form: " & Err.description, vbCritical
    Resume ErrorHandler_Exit
End Sub
```

#### GetNewCable() - Creates Cable Object with Checkbox Values

```vba
Public Function GetNewCable() As Variant
    On Error GoTo ErrorHandler

    ' CREATE NEW CABLE OBJECT
    Dim cNewCable As New clCable
    Dim strSourceDesc As String
    Dim strDestDesc As String
    
    ' Get endpoint descriptions
    strSourceDesc = modDatabase.GetEndpointDescription(ID_NEW_CAB_FORM, CStr(Me.cmb_Source.Value))
    strDestDesc = modDatabase.GetEndpointDescription(ID_NEW_CAB_FORM, CStr(Me.cmb_Destination.Value))

    ' POPULATE CABLE OBJECT WITH FORM DATA
    With cNewCable
        ' *** CHECKBOX VALUES - READ FROM FORM CONTROLS ***
        .Scheduled = Me.cb_IsScheduled.Value        ' Boolean from checkbox
        .IDAttached = Me.cb_IDAttached.Value        ' Boolean from checkbox
        
        ' Text field values (String)
        .cableID = Me.txt_CableID.Value
        .Source = strSourceDesc
        .Destination = strDestDesc
        .CableLength = Me.txt_Length.Value
        
        ' Dropdown selection values (String)
        .CoreSize = Me.cmb_CoreSize.Value
        .EarthSize = Me.cmb_EarthSize.Value
        .CoreConfig = Me.cmb_CoreConfig.Value
        .InsulationType = Me.cmb_Insulation.Value
        .CableType = Me.cmb_CableType.Value
    End With
    
    ' RETURN POPULATED CABLE OBJECT
    Set GetNewCable = cNewCable

ErrorHandler_Exit:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred creating cable object: " & Err.description, vbCritical
    Set GetNewCable = Nothing
    Resume ErrorHandler_Exit
End Function
```

#### ShowForUpdate() - Populates Checkboxes for Editing

```vba
Public Sub ShowForUpdate(cableToEdit As clCable, strFormID As String, lngRowIndex As Long)
    On Error GoTo ErrorHandler
    
    ' Set form mode to UPDATE
    FormMode = "UPDATE"
    UpdateRowIndex = lngRowIndex
    
    ' ... lookup code omitted ...
    
    ' NOW POPULATE THE FORM FIELDS
    With Me
        ' *** RESTORE CHECKBOX VALUES FROM CABLE OBJECT ***
        .cb_IsScheduled.Value = cableToEdit.Scheduled
        .cb_IDAttached.Value = cableToEdit.IDAttached
        
        ' Text fields
        .txt_CableID.Value = cableToEdit.cableID
        .txt_CircuitName.Value = ExtractCircuitName(cableToEdit.cableID)
        .txt_Length.Value = cableToEdit.CableLength
        
        ' ... other controls ...
    End With
    
    ' ... rest of method ...
ErrorHandler_Exit:
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred loading cable for edit: " & Err.description, vbCritical
    Resume ErrorHandler_Exit
End Sub
```

---

## 3. SAVE OPERATIONS - cmd_Save_Click()

```vba
Private Sub cmd_Save_Click()
    On Error GoTo ErrorHandler
    
    ' VALIDATE ALL REQUIRED FIELDS BEFORE PROCEEDING
    If Not Me.ValidateRequired() Then
        Exit Sub
    End If
    
    ' CREATE CABLE OBJECT FROM FORM DATA
    ' This calls GetNewCable() which reads checkbox values
    Dim cNewCable As New clCable
    Set cNewCable = Me.GetNewCable()
    
    If cNewCable Is Nothing Then
        MsgBox "Failed to create cable object. Please check your data and try again.", vbCritical
        Exit Sub
    End If
    
    ' SAVE OR UPDATE CABLE BASED ON FORM MODE
    Dim bSuccess As Boolean
    
    If FormMode = "UPDATE" Then
        ' UPDATE MODE: Update existing cable
        ' Checkbox values are in cNewCable object and will be written to table
        bSuccess = modDatabase.UpdateCable(ID_NEW_CAB_FORM, UpdateRowIndex, cNewCable)
        
        If bSuccess Then
            MsgBox "Cable updated successfully!", vbInformation, "Update Successful"
            Unload Me
        Else
            MsgBox "An error occurred updating the cable. Please check your data and try again.", vbCritical
        End If
        
    Else
        ' CREATE MODE: Register new cable
        bSuccess = modDatabase.RegisterCable(ID_NEW_CAB_FORM, cNewCable)
        
        If bSuccess Then
            MsgBox "New cable entered successfully!", vbInformation, "Save Successful"
            Unload Me
        Else
            MsgBox "An error occurred saving the cable. Please check your data and try again.", vbCritical
        End If
    End If

ErrorHandler_Exit:
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred during save: " & Err.description, vbCritical, "Save Error"
    Resume ErrorHandler_Exit
End Sub
```

---

## 4. DATABASE LAYER - modDatabase.bas

### RegisterCable() - New Cable Registration

```vba
Public Function RegisterCable(ByVal strFormID As String, _
                                ByVal cNewCable As clCable) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lngNewRowNumber As Long
    
    ' Route the cable save operation to the appropriate plant worksheet
    Select Case strFormID
        Case Is = "WET_PLANT": lngNewRowNumber = sht_WetPlant.SaveCable(cNewCable)
        Case Is = "ORE_SORTER": lngNewRowNumber = sht_OreSorter.SaveCable(cNewCable)
        Case Is = "RETREATMENT": lngNewRowNumber = sht_Retreatment.SaveCable(cNewCable)
    End Select
    
    ' Interpret the result: non-zero return means success
    If Not lngNewRowNumber = eErrorCodes.ecError Then
        RegisterCable = True
    Else
        RegisterCable = False
    End If
    
    Set cNewCable = Nothing

ErrorExit:
Exit Function
ErrorHandler:
    Set cNewCable = Nothing
    HandleError "modDatabase", "RegisterCable", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function
```

### UpdateCable() - Existing Cable Update

```vba
Public Function UpdateCable(strFormID As String, lngRowIndex As Long, cUpdatedCable As clCable) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bSuccess As Boolean
    
    ' Route the cable update request to the appropriate plant worksheet
    Select Case strFormID
        Case Is = "WET_PLANT"
            bSuccess = sht_WetPlant.UpdateCable(lngRowIndex, cUpdatedCable)
        Case Is = "ORE_SORTER"
            bSuccess = sht_OreSorter.UpdateCable(lngRowIndex, cUpdatedCable)
        Case Is = "RETREATMENT"
            bSuccess = sht_Retreatment.UpdateCable(lngRowIndex, cUpdatedCable)
        Case Else
            bSuccess = False
    End Select
    
    UpdateCable = bSuccess
    Set cUpdatedCable = Nothing

ErrorExit:
Exit Function
ErrorHandler:
    Set cUpdatedCable = Nothing
    HandleError "modDatabase", "UpdateCable", Err.Number, Err.description, Erl
    UpdateCable = False
    Resume ErrorExit
End Function
```

---

## 5. WORKSHEET MODULES - sht_WetPlant.cls

### SaveCable() - Writes New Cable Row

```vba
Public Function SaveCable(ByVal cNewCable As clCable) As Long
    On Error GoTo ErrorHandler
    
    ' Get cable data as array using the object's ToRow method
    ' ToRow() converts the cable object (with Boolean checkbox values) to an array
    Dim arrRow() As Variant
    arrRow = cNewCable.ToRow()
    
    ' Add new row to the table
    Dim lrNewRow As ListRow
    Set lrNewRow = Me.ListObjects("tbl_WetPlantCables").ListRows.Add
    
    ' Populate all columns with cable data
    Dim i As Long
    For i = 1 To ccTotalColumns
        ' Column 1 and 2 receive the Boolean checkbox values
        lrNewRow.Range(i) = arrRow(i)
    Next i
        
    ' Return the total number of cables (including the new one)
    Dim lngNewRowNumber As Long
    lngNewRowNumber = Me.GetNumberOfCables()
    
    SaveCable = lngNewRowNumber
    
ErrorExit:
Exit Function
ErrorHandler:
    HandleError "sht_WetPlant", "SaveCable", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function
```

### UpdateCable() - Writes Updated Cable Row

```vba
Public Function UpdateCable(lngRowIndex As Long, cUpdatedCable As clCable) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate row index exists within table bounds
    If lngRowIndex < 1 Or lngRowIndex > Me.ListObjects("tbl_WetPlantCables").ListRows.count Then
        MsgBox "Invalid row index: " & lngRowIndex, vbExclamation
        UpdateCable = False
        Exit Function
    End If
    
    ' Update the row with new cable data
    ' ToRow() converts the cable object to an array with checkbox Boolean values
    Dim arrRow As Variant
    arrRow = cUpdatedCable.ToRow()
    
    ' Write the updated data to the table row
    ' This includes the Boolean values for columns 1 (Scheduled) and 2 (IDAttached)
    Me.ListObjects("tbl_WetPlantCables").ListRows(lngRowIndex).Range.Value = arrRow
    
    ' Return success
    UpdateCable = True
    
ErrorExit:
Exit Function
ErrorHandler:
    HandleError "sht_WetPlant", "UpdateCable", Err.Number, Err.description, Erl
    UpdateCable = False
    Resume ErrorExit
End Function
```

---

## 6. DATA MAPPING - GetCablesArray()

When reading data from the table back into Cable objects:

```vba
Public Function GetCablesArray() As Variant
    On Error GoTo ErrorHandler
    
    Dim cTemp As New clCable
    Dim arrRows As Variant
    Dim arrCables() As New clCable
    
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
    
    ' Process each row of data
    For lngRow = 1 To UBound(arrRows, 1)
        ' Process each column and map to cable object property
        For lngCol = 1 To UBound(arrRows, 2)
            If lngCol = ccScheduled Then
                ' Read Boolean value from column 1 into Scheduled property
                arrCables(lngRow).Scheduled = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = ccIDAttached Then
                ' Read Boolean value from column 2 into IDAttached property
                arrCables(lngRow).IDAttached = arrRows(lngRow, lngCol)
                
            ElseIf lngCol = ccCableID Then
                arrCables(lngRow).cableID = arrRows(lngRow, lngCol)
            
            ' ... other columns ...
            End If
        Next lngCol
    Next lngRow
        
    GetCablesArray = arrCables

ErrorExit:
Exit Function
ErrorHandler:
    Set arrRows = Nothing
    HandleError "sht_WetPlant", "GetCablesArray", Err.Number, Err.description, Erl
    Resume ErrorExit
End Function
```

---

## Data Flow Diagram

```
FORM INPUT:
cb_IsScheduled (checkbox) ─┐
cb_IDAttached (checkbox)  ─┤
                           ├─> GetNewCable()
(other form controls)     ─┘
                           |
                           v
                     clCable Object
                    (mIsScheduled, mIDAttached)
                           |
                           v
                       ToRow() Method
                    (Boolean array elements)
                           |
                           v
                    SaveCable() / UpdateCable()
                           |
                           v
              Excel Table tbl_WetPlantCables
            (Columns 1 & 2 contain Boolean values)
```

---

## Key Findings

1. **Columns 1-2 are Boolean**: The table uses native Boolean data type for checkboxes
   - Column 1: `ccScheduled` (Scheduled)
   - Column 2: `ccIDAttached` (ID Attached)

2. **Data Type Path**: Boolean checkbox → clCable object property → ToRow() array → Excel table cell

3. **Save Flow**:
   - Form checkbox value → clCable Boolean property
   - clCable.ToRow() → converts to array
   - SaveCable()/UpdateCable() → writes array to table row

4. **Update Flow**:
   - Form reads from clCable object
   - ShowForUpdate() populates checkboxes with cableToEdit.Scheduled/IDAttached
   - User modifies checkboxes
   - GetNewCable() reads updated values
   - UpdateCable() writes to table

5. **Three Worksheet Modules**: Same pattern across all three plants:
   - sht_WetPlant
   - sht_OreSorter
   - sht_Retreatment

---

## Notes on Checkbox Formatting

The issue mentioned in the docs about "checkbox columns losing their checkbox formatting after updates" is likely caused by:

1. **Table style resets**: The `ResetTableStyle()` method clears and reapplies the table style
2. **Direct range assignment**: Writing to `ListRows(lngRowIndex).Range.Value = arrRow` directly assigns values rather than using cell-by-cell assignment
3. **Excel's table formatting behavior**: Direct array assignment can strip formatting

To preserve checkbox formatting, consider:
- Using cell-by-cell assignment instead of array assignment
- Explicitly applying checkbox formatting after updates
- Using the `ResetTableStyle()` method after updates


# Debugging the Save Error

## Add Debug Messages Earlier in the Flow

### Step 1: Add Debug to cmd_Save_Click - Very Start

1. Open VBA Editor (`Alt + F11`)
2. Find `frm_RegisterCable`
3. Search for `Private Sub cmd_Save_Click()`
4. Add debug at the **very beginning**:

```vba
Private Sub cmd_Save_Click()
    On Error GoTo ErrorHandler

    Debug.Print "=== SAVE CLICKED ==="
    Debug.Print "FormMode: " & FormMode
    Debug.Print "ID_NEW_CAB_FORM: " & ID_NEW_CAB_FORM
    Debug.Print "UpdateRowIndex: " & UpdateRowIndex
    Debug.Print "====================="

    ' -----------------------------------------------------------------------
    ' VALIDATE ALL REQUIRED FIELDS BEFORE PROCEEDING
    ' -----------------------------------------------------------------------
    If Not Me.ValidateRequired() Then
        Debug.Print "Validation failed!"
        Exit Sub
    End If

    Debug.Print "Validation passed, creating cable object..."

    ' -----------------------------------------------------------------------
    ' CREATE CABLE OBJECT FROM FORM DATA
    ' -----------------------------------------------------------------------
    Dim cNewCable As New clCable
    Set cNewCable = Me.GetNewCable()

    Debug.Print "GetNewCable returned: " & (Not cNewCable Is Nothing)

    ' Check if cable object creation was successful
    If cNewCable Is Nothing Then
        Debug.Print "ERROR: Cable object is Nothing!"
        MsgBox "Failed to create cable object. Please check your data and try again.", vbCritical, "Save Error"
        Exit Sub
    End If
```

### Step 2: Add Debug to GetNewCable

Search for `Public Function GetNewCable()` and add:

```vba
Public Function GetNewCable() As Variant
    On Error GoTo ErrorHandler

    Debug.Print "=== GetNewCable START ==="
    Debug.Print "ID_NEW_CAB_FORM: " & ID_NEW_CAB_FORM
    Debug.Print "cmb_Source.Value: [" & Me.cmb_Source.Value & "]"
    Debug.Print "cmb_Destination.Value: [" & Me.cmb_Destination.Value & "]"

    ' -----------------------------------------------------------------------
    ' CREATE NEW CABLE OBJECT
    ' -----------------------------------------------------------------------
    Dim cNewCable As New clCable          ' New cable object to populate
    Dim strSourceDesc As String           ' Full description of source endpoint
    Dim strDestDesc As String             ' Full description of destination endpoint

    ' -----------------------------------------------------------------------
    ' GET ENDPOINT DESCRIPTIONS FROM DATABASE
    ' Convert short codes (like "WM101") to full descriptions
    ' -----------------------------------------------------------------------
    Debug.Print "Calling GetEndpointDescription for Source..."
    strSourceDesc = modDatabase.GetEndpointDescription(ID_NEW_CAB_FORM, CStr(Me.cmb_Source.Value))
    Debug.Print "Source description: [" & strSourceDesc & "]"

    Debug.Print "Calling GetEndpointDescription for Destination..."
    strDestDesc = modDatabase.GetEndpointDescription(ID_NEW_CAB_FORM, CStr(Me.cmb_Destination.Value))
    Debug.Print "Destination description: [" & strDestDesc & "]"
```

### Step 3: Check the ErrorHandler

At the bottom of `cmd_Save_Click`, find the ErrorHandler and make sure it prints:

```vba
ErrorHandler:
    ' Handle any unexpected errors during save operation
    Debug.Print "!!! ERROR in cmd_Save_Click !!!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Line: " & Erl

    MsgBox "An error occurred during save: " & Err.Description, vbCritical, "Save Error"
    Debug.Print "Error in cmd_Save_Click: " & Err.Number & " - " & Err.Description
    Resume ErrorHandler_Exit
End Sub
```

### Step 4: Test and Report

1. Save the VBA code
2. Press `Ctrl + G` to open Immediate Window
3. Try to save the cable again
4. **Copy ALL the output from the Immediate Window and paste it here**

This will show us exactly where it's failing!

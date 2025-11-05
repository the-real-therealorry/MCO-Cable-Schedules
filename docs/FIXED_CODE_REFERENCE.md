# Fixed Code Reference

This document shows the exact code sections that need to be modified.

---

## Change 1: Module-Level Variables

**LOCATION:** Top of frm_RegisterCable, around line 20-35

### BEFORE:
```vba
' Form mode tracking - indicates whether form is in CREATE or UPDATE mode
Dim FormMode As String  ' Values: "CREATE" or "UPDATE"

' Row index for update operations - stores which row we're editing
Dim UpdateRowIndex As Long
```

### AFTER:
```vba
' Form mode tracking - indicates whether form is in CREATE or UPDATE mode
Dim FormMode As String  ' Values: "CREATE" or "UPDATE"

' Row index for update operations - stores which row we're editing
Dim UpdateRowIndex As Long

' Initialization flag - prevents UserForm_Activate from re-running
Private mbFormInitialized As Boolean
```

---

## Change 2: UserForm_Activate - Add Early Exit

**LOCATION:** UserForm_Activate subroutine, around line 690-695

### BEFORE:
```vba
Private Sub UserForm_Activate()
    On Error GoTo ErrorHandler

    ' -----------------------------------------------------------------------
    ' INITIALIZE FORM MODE IF NOT ALREADY SET
    ' Default to CREATE mode if not specified (for backwards compatibility)
    ' -----------------------------------------------------------------------
    If FormMode = "" Then
        FormMode = "CREATE"
    End If
```

### AFTER:
```vba
Private Sub UserForm_Activate()
    On Error GoTo ErrorHandler

    ' -----------------------------------------------------------------------
    ' PREVENT RE-INITIALIZATION IF FORM ALREADY INITIALIZED
    ' This prevents UpdateComboBox from clearing values set by ShowForUpdate
    ' -----------------------------------------------------------------------
    If mbFormInitialized Then
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' INITIALIZE FORM MODE IF NOT ALREADY SET
    ' Default to CREATE mode if not specified (for backwards compatibility)
    ' -----------------------------------------------------------------------
    If FormMode = "" Then
        FormMode = "CREATE"
    End If
```

---

## Change 3: UserForm_Activate - Set Flag at End

**LOCATION:** End of UserForm_Activate subroutine, around line 770-780

### BEFORE:
```vba
            MsgBox "Unknown form ID: " & ID_NEW_CAB_FORM & vbCrLf & _
                   "Please check form configuration.", vbCritical, "Configuration Error"
    End Select

ErrorHandler_Exit:
    Exit Sub

ErrorHandler:
    ' Handle any unexpected errors during form activation
    MsgBox "An error occurred activating the form: " & Err.Description, vbCritical, "Activation Error"
    Debug.Print "Error in UserForm_Activate: " & Err.Number & " - " & Err.Description
    Resume ErrorHandler_Exit
End Sub
```

### AFTER:
```vba
            MsgBox "Unknown form ID: " & ID_NEW_CAB_FORM & vbCrLf & _
                   "Please check form configuration.", vbCritical, "Configuration Error"
    End Select

    ' -----------------------------------------------------------------------
    ' MARK FORM AS INITIALIZED
    ' Prevents re-initialization on subsequent activations
    ' -----------------------------------------------------------------------
    mbFormInitialized = True

ErrorHandler_Exit:
    Exit Sub

ErrorHandler:
    ' Handle any unexpected errors during form activation
    MsgBox "An error occurred activating the form: " & Err.Description, vbCritical, "Activation Error"
    Debug.Print "Error in UserForm_Activate: " & Err.Number & " - " & Err.Description
    Resume ErrorHandler_Exit
End Sub
```

---

## Change 4: Add UserForm_Terminate

**LOCATION:** Add as new subroutine after UserForm_Activate

### ADD THIS NEW SUBROUTINE:
```vba
' ===============================================================================
' SUBROUTINE: UserForm_Terminate
' PURPOSE: Cleanup when form is destroyed
' NOTES: Resets initialization flag for next use
' ===============================================================================
Private Sub UserForm_Terminate()
    On Error Resume Next

    ' Reset initialization flag when form is destroyed
    mbFormInitialized = False
End Sub
```

---

## Change 5: ShowForUpdate - Set Flag Before Show

**LOCATION:** End of ShowForUpdate subroutine, around line 1260-1270

### BEFORE:
```vba
    ' Update form caption
    Me.Caption = "Edit Cable - " & cableToEdit.cableID

    ' Disable key fields during edit
    Me.txt_CableID.Enabled = False
    Me.txt_CircuitName.Enabled = False
    Me.cmb_Source.Enabled = False
    Me.cmb_Destination.Enabled = False

    ' Show the form
    Me.Show

ErrorHandler_Exit:
    Exit Sub
```

### AFTER:
```vba
    ' Update form caption
    Me.Caption = "Edit Cable - " & cableToEdit.cableID

    ' Disable key fields during edit
    Me.txt_CableID.Enabled = False
    Me.txt_CircuitName.Enabled = False
    Me.cmb_Source.Enabled = False
    Me.cmb_Destination.Enabled = False

    ' -----------------------------------------------------------------------
    ' SET INITIALIZATION FLAG BEFORE SHOWING FORM
    ' This prevents UserForm_Activate from clearing the combo box values
    ' -----------------------------------------------------------------------
    mbFormInitialized = True

    ' Show the form
    Me.Show

ErrorHandler_Exit:
    Exit Sub
```

---

## Summary of Changes

1. ✅ Add `mbFormInitialized` flag variable
2. ✅ Check flag at start of `UserForm_Activate` and exit if True
3. ✅ Set flag to True at end of `UserForm_Activate`
4. ✅ Add `UserForm_Terminate` to reset flag
5. ✅ Set flag to True in `ShowForUpdate` before calling `Me.Show`

These 5 changes prevent the combo boxes from being cleared during edit operations.

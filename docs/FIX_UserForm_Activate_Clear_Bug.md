# Fix: UserForm_Activate Clearing Combo Box Values

## Problem Identified

When editing a cable record:
1. `ShowForUpdate()` populates combo boxes and sets their values correctly
2. `Me.Show` is called to display the form
3. **`Me.Show` triggers `UserForm_Activate` event**
4. `UserForm_Activate` calls `UpdateComboBox()` which does:
   - `Me.cmb_Source.Clear`
   - `Me.cmb_Destination.Clear`
5. **The values that were just set are CLEARED!**
6. List is repopulated but `.Value` is not restored
7. User sees empty dropdowns

## Solution: Add Initialization Flag

Add a module-level flag to prevent `UserForm_Activate` from re-running when form is already initialized.

### Code Changes Required

#### 1. Add Module-Level Variable

In `frm_RegisterCable`, add this variable with the other module-level declarations:

```vba
' Module-level flag to prevent re-initialization
Private mbFormInitialized As Boolean
```

#### 2. Modify UserForm_Activate

Update the `UserForm_Activate` subroutine to check the flag:

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

    ' ... rest of existing UserForm_Activate code ...

    ' -----------------------------------------------------------------------
    ' MARK FORM AS INITIALIZED
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

#### 3. Reset Flag on Form Termination

Add this to ensure flag resets when form closes:

```vba
Private Sub UserForm_Terminate()
    ' Reset initialization flag when form is destroyed
    mbFormInitialized = False
End Sub
```

#### 4. Ensure ShowForUpdate Sets the Flag

At the end of `ShowForUpdate`, before `Me.Show`, add:

```vba
    ' Mark form as initialized to prevent UserForm_Activate from clearing values
    mbFormInitialized = True

    ' Show the form
    Me.Show
```

## Why This Works

1. **CREATE Mode (new cable):**
   - Form is shown via normal means
   - `UserForm_Activate` runs FIRST (flag is False)
   - Initializes everything, sets flag to True
   - Works as before

2. **UPDATE Mode (edit cable):**
   - `ShowForUpdate()` pre-initializes everything
   - Sets `mbFormInitialized = True` BEFORE showing
   - `Me.Show` triggers `UserForm_Activate`
   - `UserForm_Activate` sees flag is True and exits immediately
   - Combo box values remain intact!

## Alternative Considered

We could have moved the combo box value setting to AFTER `Me.Show`, but:
- More invasive code changes
- Harder to maintain
- User might see a flicker as values populate

The flag approach is cleaner and follows standard UserForm patterns.

## Files to Update

Apply this fix to all three spreadsheet files:
1. `Power Cable Register Rev0.xlsm`
2. `Control & Instrument cable Register Rev0.xlsm`
3. `Structured Cable Register Rev0.xlsm`

All three have identical form code structure.

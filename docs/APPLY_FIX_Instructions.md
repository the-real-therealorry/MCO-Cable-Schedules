# Step-by-Step Instructions: Apply the UserForm_Activate Fix

## Quick Summary
This fix prevents the form from clearing Source/Destination dropdown values when editing a cable.

## Time Required
About 5 minutes per file (15 minutes total for all 3 files)

---

## Instructions

Apply these changes to **each** of the three Excel files:
1. `Power Cable Register Rev0.xlsm`
2. `Control & Instrument cable Register Rev0.xlsm`
3. `Structured Cable Register Rev0.xlsm`

### Step 1: Open VBA Editor

1. Open the Excel file
2. Press `Alt + F11` to open the VBA Editor
3. In the Project Explorer (left pane), find **frm_RegisterCable**
4. Double-click to open it

---

### Step 2: Add Module-Level Variable

1. Scroll to the **top** of the code where module-level variables are declared
2. You'll see lines like:
   ```vba
   Dim ID_NEW_CAB_FORM As String
   Dim strShortNamePattern As String
   Dim strCircuitNamePattern As String
   Dim strCableIDPattern As String
   Dim count As Integer
   Dim FormMode As String
   Dim UpdateRowIndex As Long
   ```

3. **Add this new line** after `UpdateRowIndex`:
   ```vba
   Private mbFormInitialized As Boolean
   ```

4. Result should look like:
   ```vba
   Dim UpdateRowIndex As Long
   Private mbFormInitialized As Boolean
   ```

---

### Step 3: Modify UserForm_Activate

1. Press `Ctrl + F` to open Find dialog
2. Search for: `Private Sub UserForm_Activate()`
3. Click **Find Next**

4. You'll see code that starts like:
   ```vba
   Private Sub UserForm_Activate()
       On Error GoTo ErrorHandler

       ' INITIALIZE FORM MODE IF NOT ALREADY SET
       If FormMode = "" Then
           FormMode = "CREATE"
       End If
   ```

5. **Add these lines** RIGHT AFTER `On Error GoTo ErrorHandler`:
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

       ' INITIALIZE FORM MODE IF NOT ALREADY SET
       If FormMode = "" Then
           FormMode = "CREATE"
       End If
   ```

6. Scroll to the **end** of the `UserForm_Activate` subroutine (before `ErrorHandler_Exit:`)

7. **Add this line** just before the error handler section:
   ```vba
       ' Set flag to prevent re-initialization
       mbFormInitialized = True

   ErrorHandler_Exit:
       Exit Sub
   ```

---

### Step 4: Add UserForm_Terminate

1. Scroll to find `UserForm_Initialize` or another UserForm event handler
2. **Add this new subroutine** nearby:
   ```vba
   Private Sub UserForm_Terminate()
       ' Reset initialization flag when form is destroyed
       mbFormInitialized = False
   End Sub
   ```

---

### Step 5: Modify ShowForUpdate

1. Press `Ctrl + F` to search for: `Public Sub ShowForUpdate`
2. Scroll to the **end** of this subroutine
3. Find these lines near the end:
   ```vba
       ' Show the form
       Me.Show
   ```

4. **Add one line BEFORE `Me.Show`**:
   ```vba
       ' Mark form as initialized to prevent UserForm_Activate from clearing values
       mbFormInitialized = True

       ' Show the form
       Me.Show
   ```

---

### Step 6: Save and Test

1. Press `Ctrl + S` to save
2. Close the VBA Editor
3. Save the Excel file
4. **Test it:**
   - Create a new cable record with Source and Destination
   - Click the Edit button on that cable
   - **Source and Destination should now be populated!**

---

## Verification

After applying the fix, when you edit a cable:
- ✅ Source dropdown should show the correct value
- ✅ Destination dropdown should show the correct value
- ✅ No more data loss when saving

If you still see empty dropdowns, double-check that you added all 4 changes in the correct locations.

---

## Need Help?

If you encounter issues:
1. Check the VBA Immediate Window (`Ctrl + G`) for error messages
2. Verify all 4 code changes were applied
3. Make sure you saved the file after editing
4. Try closing and reopening Excel

---

## Rollback

If something goes wrong, you can restore from git:
```bash
git checkout HEAD -- "Power Cable Register Rev0.xlsm"
```

(But don't commit your data - only commit the template!)

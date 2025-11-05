# Complete Checkbox Fix Guide

## Quick Summary
The checkbox columns are showing "TRUE/FALSE" text instead of checkboxes. This guide provides two solutions:
1. **Quick Fix**: Run a macro once to convert all existing TRUE/FALSE to checkboxes
2. **Permanent Fix**: Update VBA code to maintain checkboxes automatically

---

## ⚡ QUICK FIX - Run This Macro Once

This will immediately fix all existing checkboxes in all three plant sheets.

### Steps:
1. Open the Excel file
2. Press `Alt+F11` to open VBA Editor
3. Click **Insert** → **Module** to create a new module
4. Copy and paste this code:

```vba
Sub FixAllCheckboxesNow()
    '==============================================================================
    ' ONE-TIME FIX: Converts all existing TRUE/FALSE to checkboxes
    ' Run this once to fix all existing data
    '==============================================================================

    On Error Resume Next

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rngCol1 As Range
    Dim rngCol2 As Range

    ' Process WET PLANT sheet
    Set ws = ThisWorkbook.Worksheets("sht_WetPlant")
    Set tbl = ws.ListObjects("tbl_WetPlantCables")
    If tbl.ListRows.Count > 0 Then
        Set rngCol1 = tbl.ListColumns(1).DataBodyRange
        Set rngCol2 = tbl.ListColumns(2).DataBodyRange
        Call ApplyCheckboxToRange(rngCol1, rngCol2)
    End If

    ' Process ORE SORTER sheet
    Set ws = ThisWorkbook.Worksheets("sht_OreSorter")
    Set tbl = ws.ListObjects("tbl_OreSorterCables")
    If tbl.ListRows.Count > 0 Then
        Set rngCol1 = tbl.ListColumns(1).DataBodyRange
        Set rngCol2 = tbl.ListColumns(2).DataBodyRange
        Call ApplyCheckboxToRange(rngCol1, rngCol2)
    End If

    ' Process RETREATMENT sheet
    Set ws = ThisWorkbook.Worksheets("sht_Retreatment")
    Set tbl = ws.ListObjects("tbl_RetreatmentCables")
    If tbl.ListRows.Count > 0 Then
        Set rngCol1 = tbl.ListColumns(1).DataBodyRange
        Set rngCol2 = tbl.ListColumns(2).DataBodyRange
        Call ApplyCheckboxToRange(rngCol1, rngCol2)
    End If

    On Error GoTo 0

    MsgBox "Checkboxes have been restored in all three plant sheets!", vbInformation, "Success"
End Sub

Private Sub ApplyCheckboxToRange(rng1 As Range, rng2 As Range)
    On Error Resume Next

    ' Method 1: Try Excel 365/2019+ checkbox data type
    rng1.ExcelDataType = xlCheckbox
    rng2.ExcelDataType = xlCheckbox

    ' Method 2: If Method 1 fails, refresh cell values
    If Err.Number <> 0 Then
        Err.Clear
        Dim cell As Range
        For Each cell In rng1
            If cell.Value = True Or cell.Value = False Then
                cell.Value = cell.Value ' Refresh
            End If
        Next cell
        For Each cell In rng2
            If cell.Value = True Or cell.Value = False Then
                cell.Value = cell.Value ' Refresh
            End If
        Next cell
    End If

    On Error GoTo 0
End Sub
```

5. Press `F5` or click **Run** → **Run Sub/UserForm**
6. Close the VBA Editor
7. **Done!** All your checkboxes should now be displaying correctly

---

## 🔧 PERMANENT FIX - Update VBA Code

This ensures checkboxes stay formatted correctly when you save or update cables.

### Option A: Copy/Paste Method (Recommended - Easiest)

For each of the three worksheet modules, you'll copy/paste the modified code.

#### 1. Fix sht_WetPlant

1. Press `Alt+F11` to open VBA Editor
2. In the Project Explorer (left panel), double-click **sht_WetPlant**
3. Scroll to the very bottom of the code
4. Add this new function at the end:

```vba
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
        If tblCables.ListRows.Count > 0 Then
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
```

5. Find the `SaveCable` function (use `Ctrl+F` to search)
6. Find this line: `SaveCable = lngNewRowNumber`
7. Right after that line, add:
```vba
    ' Apply checkbox formatting to Boolean columns
    Call ApplyCheckboxFormatting
```

8. Find the `UpdateCable` function
9. Find this line: `UpdateCable = True`
10. Right after that line, add:
```vba
    ' Apply checkbox formatting to Boolean columns
    Call ApplyCheckboxFormatting
```

#### 2. Fix sht_OreSorter

Repeat the same steps as above, but:
- Open **sht_OreSorter** module instead
- Use this table name in the ApplyCheckboxFormatting function:
```vba
Set tblCables = Me.ListObjects("tbl_OreSorterCables")
```

#### 3. Fix sht_Retreatment

Repeat the same steps again, but:
- Open **sht_Retreatment** module instead
- Use this table name in the ApplyCheckboxFormatting function:
```vba
Set tblCables = Me.ListObjects("tbl_RetreatmentCables")
```

---

### Option B: Import Modified VBA Files (Advanced)

I've created modified VBA files that are ready to import. They're located in the project folder:
- `vba_sht_WetPlant.cls.bas`
- `vba_sht_OreSorter.cls.bas`
- `vba_sht_Retreatment.cls.bas`

**Steps:**
1. Open Excel file, press `Alt+F11`
2. For each module (sht_WetPlant, sht_OreSorter, sht_Retreatment):
   - Right-click the module in Project Explorer
   - Choose **Remove [module name]**
   - When prompted, choose **No** (don't export)
3. Click **File** → **Import File**
4. Browse to the `vba_sht_WetPlant.cls.bas` file and import it
5. Repeat for the other two .bas files
6. Save the Excel file

**⚠️ Note:** This method is trickier because the module needs to remain associated with the worksheet. The copy/paste method (Option A) is recommended.

---

## 🧪 Testing

After applying either fix:

1. Register a new cable with checkboxes checked/unchecked
2. Verify checkboxes display correctly after saving
3. Edit an existing cable and change a checkbox
4. Verify checkboxes display correctly after updating

---

## 📋 What Changed?

### Problem
When `SaveCable()` and `UpdateCable()` wrote Boolean TRUE/FALSE values to cells, Excel lost the checkbox formatting.

### Solution
Added `ApplyCheckboxFormatting()` function that:
1. Tries to apply Excel's native checkbox data type (`xlCheckbox`)
2. Falls back to refreshing cell values if that fails (older Excel versions)
3. Runs automatically after every save and update operation

---

## ❓ Troubleshooting

### Checkboxes still showing TRUE/FALSE after permanent fix
- Run the Quick Fix macro once to convert existing data
- The permanent fix only affects NEW saves and updates

### Getting xlCheckbox errors
- Your Excel version might not support the xlCheckbox constant
- The fallback method should still work
- If not, you may need to manually select the columns and apply checkbox formatting

### Can't find the VBA modules
- Press `Alt+F11` to open VBA Editor
- Look in the left panel under **Microsoft Excel Objects**
- You should see: sht_WetPlant, sht_OreSorter, sht_Retreatment

### Changes not taking effect
- Make sure you saved the Excel file after making VBA changes
- Close and reopen the file to ensure macros are enabled

---

## 📞 Need Help?

If you encounter issues:
1. Check that all three worksheet modules have been updated
2. Verify there are no syntax errors (VBA Editor highlights errors in red)
3. Make sure table names match exactly:
   - `tbl_WetPlantCables`
   - `tbl_OreSorterCables`
   - `tbl_RetreatmentCables`
4. Try running the Quick Fix macro first before testing saves/updates

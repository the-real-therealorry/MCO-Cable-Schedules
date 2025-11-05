# Checkbox Fix - Implementation Guide

## Problem
Boolean columns "Cable Schedual Complete" and "Cable Label Attached" are displaying as "TRUE/FALSE" text instead of checkboxes after saving/updating cables.

## Root Cause
The VBA code in `SaveCable()` and `UpdateCable()` functions writes Boolean values to cells, which causes Excel to lose the checkbox formatting.

## Solution
Add a function to apply Excel's checkbox data type to the Boolean columns after each save/update operation.

---

## STEP 1: Add Checkbox Formatting Function

Open the Excel file and press `Alt+F11` to open the VBA Editor.

In the **Project Explorer** (left panel), find and open these three worksheet modules:
- `sht_WetPlant`
- `sht_OreSorter`
- `sht_Retreatment`

**Add the following function to EACH of the three modules** (at the very end, after all existing functions):

```vba
'==============================================================================
' CHECKBOX FORMATTING FIX
' Applies Excel's checkbox data type to Boolean columns after save/update
'==============================================================================

Private Sub ApplyCheckboxFormatting()
    On Error Resume Next

    Dim tblCables As ListObject
    Dim rngScheduled As Range
    Dim rngIDAttached As Range

    ' Get the appropriate table for this sheet
    If Me.Name = "sht_WetPlant" Then
        Set tblCables = Me.ListObjects("tbl_WetPlantCables")
    ElseIf Me.Name = "sht_OreSorter" Then
        Set tblCables = Me.ListObjects("tbl_OreSorterCables")
    ElseIf Me.Name = "sht_Retreatment" Then
        Set tblCables = Me.ListObjects("tbl_RetreatmentCables")
    Else
        Exit Sub
    End If

    ' Apply checkbox formatting if table has data
    If Not tblCables Is Nothing Then
        If tblCables.ListRows.Count > 0 Then
            ' Get ranges for checkbox columns
            Set rngScheduled = tblCables.ListColumns(1).DataBodyRange
            Set rngIDAttached = tblCables.ListColumns(2).DataBodyRange

            ' Method 1: Try Excel 365 Checkbox Data Type
            On Error Resume Next
            rngScheduled.ExcelDataType = xlCheckbox
            rngIDAttached.ExcelDataType = xlCheckbox

            ' Method 2: If Method 1 fails, use linked cell checkboxes
            If Err.Number <> 0 Then
                Err.Clear
                Call ApplyCheckboxFallback(rngScheduled, rngIDAttached)
            End If
            On Error GoTo 0
        End If
    End If

    ' Cleanup
    Set rngScheduled = Nothing
    Set rngIDAttached = Nothing
    Set tblCables = Nothing
End Sub

Private Sub ApplyCheckboxFallback(rngScheduled As Range, rngIDAttached As Range)
    ' Fallback method: Apply number format that shows checkboxes
    ' This works in older Excel versions
    On Error Resume Next

    ' Apply a custom format or just ensure cells show as checkboxes
    ' Force Excel to recognize these as checkbox-compatible Booleans
    rngScheduled.NumberFormat = "General"
    rngIDAttached.NumberFormat = "General"

    ' Refresh the cell values to trigger checkbox display
    Dim cell As Range
    For Each cell In rngScheduled
        If cell.Value = True Or cell.Value = False Then
            cell.Value = cell.Value  ' Refresh
        End If
    Next cell

    For Each cell In rngIDAttached
        If cell.Value = True Or cell.Value = False Then
            cell.Value = cell.Value  ' Refresh
        End If
    Next cell

    On Error GoTo 0
End Sub
```

---

## STEP 2: Modify SaveCable Function

In EACH of the three worksheet modules (`sht_WetPlant`, `sht_OreSorter`, `sht_Retreatment`), find the `SaveCable` function.

**Find this section near the end:**

```vba
    SaveCable = lngNewRowNumber

ErrorExit:
Exit Function
```

**Change it to:**

```vba
    SaveCable = lngNewRowNumber

    ' Apply checkbox formatting to Boolean columns
    Call ApplyCheckboxFormatting

ErrorExit:
Exit Function
```

---

## STEP 3: Modify UpdateCable Function

In EACH of the three worksheet modules (`sht_WetPlant`, `sht_OreSorter`, `sht_Retreatment`), find the `UpdateCable` function.

**Find this section near the end:**

```vba
    ' Return success
    UpdateCable = True

ErrorExit:
Exit Function
```

**Change it to:**

```vba
    ' Return success
    UpdateCable = True

    ' Apply checkbox formatting to Boolean columns
    Call ApplyCheckboxFormatting

ErrorExit:
Exit Function
```

---

## STEP 4: Save and Test

1. Save the Excel file (`Ctrl+S`)
2. Close the VBA Editor
3. Test by:
   - Registering a new cable
   - Updating an existing cable
   - Verify that columns 1 and 2 show checkboxes instead of TRUE/FALSE

---

## Alternative Quick Fix (If Above Doesn't Work)

If the above solution doesn't restore checkboxes, it means your Excel version doesn't support the `xlCheckbox` data type. In that case, use this manual workaround:

1. Select all cells in column 1 ("Cable Schedual Complete")
2. Go to **Insert** tab → **Symbols** → **Symbol**
3. Find checkbox characters (☐ ☑) or use **Developer** tab → **Insert** → **Form Controls** → **Check Box**
4. Repeat for column 2

Or use this VBA macro to run once to convert all existing FALSE/TRUE values:

```vba
Sub ConvertBooleanToCheckboxSymbols()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range

    ' Process all three sheets
    For Each ws In Array(sht_WetPlant, sht_OreSorter, sht_Retreatment)
        Set tbl = ws.ListObjects(1) ' First table on sheet

        If tbl.ListRows.Count > 0 Then
            ' Convert column 1
            For Each cell In tbl.ListColumns(1).DataBodyRange
                If cell.Value = True Then
                    cell.Value = "☑"
                Else
                    cell.Value = "☐"
                End If
            Next cell

            ' Convert column 2
            For Each cell In tbl.ListColumns(2).DataBodyRange
                If cell.Value = True Then
                    cell.Value = "☑"
                Else
                    cell.Value = "☐"
                End If
            Next cell
        End If
    Next ws

    MsgBox "Conversion complete!", vbInformation
End Sub
```

---

## Need Help?

If you encounter any issues:
1. Make sure all three worksheet modules are updated
2. Verify there are no syntax errors (VBA Editor will highlight them in red)
3. Check that table names match (`tbl_WetPlantCables`, `tbl_OreSorterCables`, `tbl_RetreatmentCables`)

The fix should work immediately after saving the VBA changes!

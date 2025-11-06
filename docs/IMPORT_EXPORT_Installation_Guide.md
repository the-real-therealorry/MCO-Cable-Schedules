# Import/Export Feature - Installation Guide

## What's Been Created

I've built a complete import/export system with three VBA modules:

1. **modImportExport** - Core export/import functions (CSV & JSON)
2. **modCompatibilityFix** - Auto-fix for missing endpoints and data issues
3. **modBackup** - Automatic backup with retention (keeps last 10)

## Installation Steps

### Step 1: Import VBA Modules

1. Open your Excel file (`Power Cable Register Rev0.xlsm`)
2. Press `Alt + F11` to open VBA Editor
3. Go to **File → Import File**
4. Navigate to `vba_code/` folder and import:
   - `modImportExport.bas`
   - `modImportExport_Import.bas` (contains import functions)
   - `modCompatibilityFix.bas`
   - `modBackup.bas`

**Important**: The file `modImportExport_Import.bas` contains additional functions that need to be **copied and pasted** into the `modImportExport` module (they're in a separate file because the module was too long to create as one file).

### Step 2: Merge Import Functions

1. Open `modImportExport_Import.bas` in a text editor
2. Copy ALL the code (Ctrl+A, Ctrl+C)
3. In VBA Editor, open `modImportExport` module
4. Scroll to the bottom
5. Paste the import functions (Ctrl+V)
6. Save
7. You can now delete the `modImportExport_Import` module from the project (it's been merged)

### Step 3: Create Required Folders

Create these folders in the same directory as your Excel file:

```
Project Root/
├── Power Cable Register Rev0.xlsm
├── Exports/        (create this folder)
├── Backups/        (will be auto-created)
└── Logs/           (create this folder)
```

### Step 4: Test Export

Let's test the export functionality:

1. Press `Alt + F11` to open VBA Editor
2. Press `Ctrl + G` to open Immediate Window
3. Type and press Enter:

```vba
? modImportExport.ExportCablesToCSV("ALL", "C:\Temp\test_cables.csv")
```

If it returns `True`, export worked! Check `C:\Temp\` for the CSV file.

### Step 5: Create Simple UI Form (Optional)

You can create a simple form to call these functions. Here's a minimal example:

1. In VBA Editor: **Insert → UserForm**
2. Name it `frm_DataManagement`
3. Add these controls:
   - Button: `cmd_ExportCSV` with caption "Export to CSV"
   - Button: `cmd_ExportJSON` with caption "Export to JSON"
   - Button: `cmd_Import` with caption "Import from File"
   - Label: `lbl_Status` for showing results

4. Add this code to the form:

```vba
Private Sub cmd_ExportCSV_Click()
    Dim filePath As String

    ' Get save location from user
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="Cables_Export_" & Format(Now, "yyyymmdd_hhnnss") & ".csv", _
        FileFilter:="CSV Files (*.csv), *.csv")

    If filePath = "False" Then Exit Sub ' User cancelled

    ' Export cables
    If modImportExport.ExportCablesToCSV("ALL", filePath) Then
        ' Export endpoints too
        Dim epPath As String
        epPath = Replace(filePath, "Cables_", "Endpoints_")
        modImportExport.ExportEndpointsToCSV "ALL", epPath

        MsgBox "Export complete!" & vbCrLf & _
               "Cables: " & filePath & vbCrLf & _
               "Endpoints: " & epPath, vbInformation
    Else
        MsgBox "Export failed. Check Immediate Window for errors.", vbCritical
    End If
End Sub

Private Sub cmd_ExportJSON_Click()
    Dim filePath As String

    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="CableRegister_Export_" & Format(Now, "yyyymmdd_hhnnss") & ".json", _
        FileFilter:="JSON Files (*.json), *.json")

    If filePath = "False" Then Exit Sub

    If modImportExport.ExportToJSON("ALL", filePath) Then
        MsgBox "JSON export complete: " & filePath, vbInformation
    Else
        MsgBox "Export failed. Check Immediate Window for errors.", vbCritical
    End If
End Sub

Private Sub cmd_Import_Click()
    Dim filePath As String
    Dim importMode As String
    Dim results As Object

    ' Ask user for import mode
    Dim response As VbMsgBoxResult
    response = MsgBox("Import Mode:" & vbCrLf & vbCrLf & _
                     "YES = Append (add to existing)" & vbCrLf & _
                     "NO = Replace (clear and import)" & vbCrLf & _
                     "CANCEL = Merge (update by ID)", _
                     vbYesNoCancel + vbQuestion, "Select Import Mode")

    Select Case response
        Case vbYes: importMode = "APPEND"
        Case vbNo: importMode = "REPLACE"
        Case vbCancel: importMode = "MERGE"
        Case Else: Exit Sub
    End Select

    ' Get file from user
    filePath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv), *.csv, JSON Files (*.json), *.json", _
        Title:="Select Import File")

    If filePath = "False" Then Exit Sub

    ' Create backup first
    Dim backupPath As String
    backupPath = modBackup.CreateBackup("ALL")

    If backupPath <> "" Then
        MsgBox "Backup created: " & vbCrLf & backupPath, vbInformation
    End If

    ' Initialize auto-fix
    modCompatibilityFix.InitializeAutoFix

    ' Import cables
    If LCase(Right(filePath, 4)) = ".csv" Then
        Set results = modImportExport.ImportCablesFromCSV(filePath, importMode)
    Else
        MsgBox "JSON import not yet implemented in this simple form.", vbInformation
        Exit Sub
    End If

    ' Show results
    If results("Success") Then
        Dim msg As String
        msg = "Import Complete!" & vbCrLf & vbCrLf
        msg = msg & "Cables Imported: " & results("CablesImported") & vbCrLf
        msg = msg & "Cables Skipped: " & results("CablesSkipped") & vbCrLf & vbCrLf
        msg = msg & modCompatibilityFix.GetAutoFixReport()

        MsgBox msg, vbInformation
    Else
        MsgBox "Import failed: " & results("ErrorMessage"), vbCritical
    End If
End Sub
```

### Step 6: Add Dashboard Button

1. Go to the Dashboard sheet
2. Insert a button (Developer tab → Insert → Button)
3. Draw the button and assign macro: `frm_DataManagement.Show`
4. Change button text to "Data Management"

---

## Usage Examples

### Export All Cables to CSV

```vba
Sub ExportAllToCSV()
    Dim timestamp As String
    Dim cablesPath As String
    Dim endpointsPath As String

    timestamp = Format(Now, "yyyymmdd_hhnnss")

    cablesPath = ThisWorkbook.Path & "\Exports\Cables_Export_" & timestamp & ".csv"
    endpointsPath = ThisWorkbook.Path & "\Exports\Endpoints_Export_" & timestamp & ".csv"

    ' Export
    If modImportExport.ExportCablesToCSV("ALL", cablesPath) Then
        modImportExport.ExportEndpointsToCSV "ALL", endpointsPath
        MsgBox "Export complete!" & vbCrLf & cablesPath, vbInformation
    End If
End Sub
```

### Export Single Plant to JSON

```vba
Sub ExportWetPlantToJSON()
    Dim filePath As String

    filePath = ThisWorkbook.Path & "\Exports\WetPlant_" & Format(Now, "yyyymmdd_hhnnss") & ".json"

    If modImportExport.ExportToJSON("WET_PLANT", filePath) Then
        MsgBox "Wet Plant exported: " & filePath, vbInformation
    End If
End Sub
```

### Import with Auto-Fix

```vba
Sub ImportWithAutoFix()
    Dim filePath As String
    Dim results As Object

    ' Create backup first
    modBackup.CreateBackup "ALL"

    ' Initialize auto-fix system
    modCompatibilityFix.InitializeAutoFix

    ' Import
    filePath = "C:\Temp\Cables_Export_20241221_103000.csv"
    Set results = modImportExport.ImportCablesFromCSV(filePath, "MERGE")

    ' Check results
    If results("Success") Then
        Debug.Print "Imported: " & results("CablesImported")
        Debug.Print "Skipped: " & results("CablesSkipped")
        Debug.Print modCompatibilityFix.GetAutoFixReport()
    Else
        Debug.Print "Error: " & results("ErrorMessage")
    End If
End Sub
```

---

## Testing Checklist

Before using in production, test these scenarios:

- [ ] Export all cables to CSV
- [ ] Export all endpoints to CSV
- [ ] Export all to JSON
- [ ] Export single plant to CSV
- [ ] Import CSV in APPEND mode
- [ ] Import CSV in REPLACE mode
- [ ] Import CSV in MERGE mode
- [ ] Verify backup files are created
- [ ] Verify old backups are auto-deleted (after 11th backup)
- [ ] Test auto-fix creates missing endpoints
- [ ] Import/export with Control & Instrument register
- [ ] Import/export with Structured Cable register

---

## Troubleshooting

### Error: "Sub or Function not defined"

**Cause**: Modules not imported correctly

**Fix**:
1. Check all 3 modules are in the project
2. Verify modImportExport_Import functions were copied into modImportExport
3. Save and close/reopen Excel

### Error: "File not found"

**Cause**: Export/Backup folders don't exist

**Fix**: Create folders manually:
- `Exports/`
- `Backups/` (will auto-create but you can create manually)
- `Logs/`

### Export creates empty files

**Cause**: Table names might be different

**Fix**: Check these table names exist in your workbook:
- `tbl_WetPlantCables`
- `tbl_OreSorterCables`
- `tbl_RetreatmentCables`
- `tbl_WetPlantEndpoints`
- `tbl_OreSorterEndpoints`
- `tbl_RetreatmentEndpoints`

### Import fails silently

**Cause**: CSV format issue

**Fix**:
1. Open CSV in Notepad (not Excel)
2. Check first line is header
3. Check commas aren't inside quotes incorrectly
4. Use JSON format instead (more robust)

---

## Advanced: Scheduled Backups

You can schedule automatic exports using Windows Task Scheduler:

1. Create a VBA macro that calls export functions
2. Save as Excel macro-enabled file
3. Create a batch file that opens Excel and runs the macro
4. Schedule the batch file in Task Scheduler

Example batch file:
```batch
"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" "C:\Path\To\YourFile.xlsm" /x
```

---

## Next Steps

1. Install all modules
2. Test export functionality
3. Create test import file
4. Test import with small dataset
5. Verify auto-fix works
6. Deploy to production

---

## Support

If you encounter issues:

1. Check Immediate Window (`Ctrl+G`) for error messages
2. Enable "Break on All Errors" in VBA (Tools → Options → General)
3. Review the code comments for function usage
4. Check the implementation plan for detailed specifications

All modules include extensive error handling and debug output!

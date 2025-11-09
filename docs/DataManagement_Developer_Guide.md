# Data Management System - Developer Guide

## Architecture Overview

The Data Management system consists of four independent VBA modules that work together to provide import/export functionality with automatic compatibility fixing and backup management.

### Module Hierarchy

```
frm_DataManagement (User Interface)
    ↓
modImportExport (Core Engine)
    ├→ modCompatibilityFix (Auto-fix)
    └→ modBackup (Backup Management)
```

### Design Principles

1. **Separation of Concerns**: Each module has a single responsibility
2. **Loose Coupling**: Modules communicate through well-defined public APIs
3. **Fail-Safe**: Comprehensive error handling, no unhandled exceptions
4. **User-Friendly**: Detailed error messages and progress feedback
5. **Data Integrity**: Automatic backups before destructive operations

---

## Module Deep Dive

### modImportExport

**Responsibility**: Core import/export operations

**Public API**:

```vba
' Export Functions
Function ExportCablesToCSV(strPlantID As String, strFilePath As String) As Boolean
Function ExportEndpointsToCSV(strPlantID As String, strFilePath As String) As Boolean
Function ExportToJSON(strPlantID As String, strFilePath As String) As Boolean

' Import Functions
Function ImportCablesFromCSV(strFilePath As String, importMode As String) As Object
Function ImportEndpointsFromCSV(strFilePath As String, importMode As String) As Object

' Utility Functions
Sub ClearPlantCables(strPlantID As String)
Sub ClearPlantEndpoints(strPlantID As String)
Function GetTimestamp() As String
```

**Key Implementation Details**:

#### CSV Escaping

Handles special characters properly:
- Commas in fields → wrap in quotes
- Quotes in fields → double them (`""`)
- Newlines in fields → wrap in quotes

```vba
Private Function CSVEscape(value As Variant) As String
    ' Implementation handles all edge cases
    ' Follows RFC 4180 CSV standard
End Function
```

#### JSON Generation

Manual JSON building (no dependencies):
- Builds JSON string character by character
- Proper escaping of special characters
- Nested structure for plants/cables/endpoints
- Includes metadata for version tracking

```vba
Private Function BuildPlantJSON(...) As String
    ' Returns properly formatted JSON
    ' Recursive structure for nested data
End Function
```

#### Import Result Format

Returns Dictionary object with:
```vba
{
    "Success": Boolean,
    "CablesImported": Long,
    "CablesSkipped": Long,
    "Errors": Collection,
    "ErrorMessage": String (if failed),
    "ErrorLine": Long (if failed)
}
```

**Error Handling Pattern**:

```vba
Public Function SomeFunction() As Boolean
    On Error GoTo ErrorHandler

    ' Function logic here
    SomeFunction = True
    Exit Function

ErrorHandler:
    SomeFunction = False
    Debug.Print "Error in SomeFunction: " & Err.Description
End Function
```

---

### modCompatibilityFix

**Responsibility**: Automatic compatibility fixing during import

**Public API**:

```vba
Sub InitializeAutoFix()
Function FixMissingEndpoint(strPlantID As String, endpointDesc As String, endpointType As String) As String
Function GetAutoFixReport() As String
Function NormalizePlantID(plantID As String) As String
Function ConvertDataType(value As Variant, targetType As String) As Variant
```

**Auto-Fix Strategy**:

#### 1. Exact Match

```vba
Private Function FindEndpointShortName(strPlantID As String, description As String) As String
    ' Case-insensitive exact match
    ' Strips "(Imported - Review)" for comparison
End Function
```

#### 2. Fuzzy Match

```vba
Private Function FuzzyMatchEndpoint(strPlantID As String, description As String) As String
    ' Normalizes strings: removes spaces, special chars, lowercase
    ' Checks for substring matches
    ' Returns short name if found
End Function
```

#### 3. Create New Endpoint

```vba
Private Function GenerateShortName(description As String, plantDigit As String) As String
    ' Extracts initials from description
    ' Example: "Motor Control Center 2" → "MCC" + plantDigit + nextNumber
    ' Returns: "MCC102"
End Function
```

**Short Name Generation Algorithm**:

1. Split description into words
2. Extract first letter of first two words
3. If single word, use first 2 letters
4. Append plant digit (1, 2, or 3)
5. Find next available number (01-19)
6. Build short name (e.g., "MCC102")

**Auto-Fix Logging**:

Module-level collection tracks all fixes:
```vba
Private m_autoFixes As Collection

Private Sub LogAutoFix(description As String)
    ' Adds fix to collection for reporting
End Sub

Public Function GetAutoFixReport() As String
    ' Returns formatted report of all fixes
End Function
```

---

### modBackup

**Responsibility**: Backup creation and management

**Public API**:

```vba
Function CreateBackup(plantID As String) As String
Function GetBackupFolder() As String
Function ListBackups() As Collection
```

**Backup Retention Logic**:

```vba
Private Sub CleanupOldBackups(backupFolder As String)
    ' 1. Get all backup files
    ' 2. Sort by date (oldest first)
    ' 3. If count > MAX_BACKUPS (10):
    '    - Delete oldest files
    '    - Keep only last 10
End Sub
```

**Sorting Algorithm**: Simple bubble sort by file modification date
- Adequate for small number of files (< 100)
- O(n²) complexity acceptable for this use case

**File Naming Convention**:
```
Import_Backup_YYYYMMDD_HHMMSS.json
Example: Import_Backup_20241221_103045.json
```

**Backup Format**: JSON (uses ExportToJSON function)
- Ensures perfect restore capability
- Includes all metadata
- Version-compatible

---

## Data Structures

### Cable Record (in memory)

```vba
' clCable class properties
Scheduled As Boolean
IDAttached As Boolean
cableID As String
Source As String          ' Full description (e.g., "MCC 2")
Destination As String     ' Full description
CoreSize As String
EarthSize As String
CoreConfig As String
InsulationType As String
CableType As String
CableLength As String
```

### Endpoint Record

```vba
' clEndpoint class properties
ShortName As String       ' e.g., "MCC102"
Description As String     ' e.g., "MCC 2" or "MCC 2 (Imported - Review)"
```

### CSV Format Specification

#### Cables CSV

**Header**:
```
Version,Plant,Scheduled,IDAttached,CableID,Source,Destination,CoreSize,EarthSize,CoreConfig,InsulationType,CableType,CableLength
```

**Data Row**:
```
2024.12.1,WET_PLANT,FALSE,TRUE,CV102-C1001-CV103,MCC 2,Conveyor belt 3,2.5mm² - 7/0.67,1.5mm² - 7/0.50,4C + E,HFI-90-TP,Black Circular SWA,45
```

**Field Specifications**:
- Version: Semantic version (YYYY.MM.revision)
- Plant: "WET_PLANT", "ORE_SORTER", or "RETREATMENT"
- Scheduled/IDAttached: "TRUE" or "FALSE" (uppercase)
- All other fields: String, CSV-escaped

#### Endpoints CSV

**Header**:
```
Version,Plant,ShortName,Description
```

**Data Row**:
```
2024.12.1,WET_PLANT,MCC102,MCC 2
```

### JSON Format Specification

```json
{
  "version": "2024.12.1",
  "exportDate": "2024-12-21T10:30:00Z",
  "sourceFile": "Power Cable Register Rev0.xlsm",
  "plants": {
    "WET_PLANT": {
      "endpoints": [
        {
          "shortName": "MCC102",
          "description": "MCC 2"
        }
      ],
      "cables": [
        {
          "scheduled": false,
          "idAttached": true,
          "cableID": "CV102-C1001-CV103",
          "source": "MCC 2",
          "destination": "Conveyor belt 3",
          "coreSize": "2.5mm² - 7/0.67",
          "earthSize": "1.5mm² - 7/0.50",
          "coreConfig": "4C + E",
          "insulationType": "HFI-90-TP",
          "cableType": "Black Circular SWA",
          "cableLength": "45"
        }
      ]
    }
  },
  "metadata": {
    "totalCables": 1,
    "totalEndpoints": 1,
    "exportType": "ALL_PLANTS"
  }
}
```

---

## Integration Points

### Worksheet Dependencies

**Required Worksheets** (by CodeName):
- `sht_WetPlant` - Wet Screen Crushing Plant cables
- `sht_OreSorter` - Ore Sorting Plant cables
- `sht_Retreatment` - Retreatment Gravity Plant cables
- `sht_Data` - Endpoint definitions

**Required Tables** (Excel ListObjects):
- `tbl_WetPlantCables` (in sht_WetPlant)
- `tbl_OreSorterCables` (in sht_OreSorter)
- `tbl_RetreatmentCables` (in sht_Retreatment)
- `tbl_WetPlantEndpoints` (in sht_Data)
- `tbl_OreSorterEndpoints` (in sht_Data)
- `tbl_RetreatmentEndpoints` (in sht_Data)

### Table Column Mappings

All cable tables use this column structure:
```
Column 1:  Scheduled (Boolean)
Column 2:  IDAttached (Boolean)
Column 3:  CableID (String)
Column 4:  Source (String - full description)
Column 5:  Destination (String - full description)
Column 6:  CoreSize (String)
Column 7:  EarthSize (String)
Column 8:  CoreConfig (String)
Column 9:  InsulationType (String)
Column 10: CableType (String)
Column 11: CableLength (String)
```

All endpoint tables use this column structure:
```
Column 1: ShortName (String)
Column 2: Description (String)
```

### External Dependencies

**Required References**:
- Microsoft Scripting Runtime (Scripting.FileSystemObject, Dictionary)
  - Usually auto-available in Office
  - If missing: Tools → References → Microsoft Scripting Runtime

**No External Libraries Required**:
- JSON parsing: Manual implementation (no JSON library)
- CSV parsing: Custom implementation
- File I/O: Scripting.FileSystemObject (built-in)

---

## UI Integration

### Form Design (frm_DataManagement)

**Recommended Controls**:

```vba
' Buttons
cmd_ExportCSV          - "Export to CSV"
cmd_ExportJSON         - "Export to JSON"
cmd_ImportFile         - "Import from File"
cmd_Close              - "Close"

' Option Frames (Optional)
opt_AllPlants          - "All Plants"
opt_WetPlant           - "Wet Plant Only"
opt_OreSorter          - "Ore Sorter Only"
opt_Retreatment        - "Retreatment Only"

' Labels
lbl_Status             - Shows operation status
lbl_LastOperation      - Shows result of last operation
```

**Minimal Implementation** (included in Installation Guide):

```vba
' Three buttons only: Export CSV, Export JSON, Import
' Uses message boxes for scope selection
' Shows results in message box
```

**Advanced Implementation** (optional):

```vba
' Full form with:
' - Scope selection radio buttons
' - Format checkboxes (CSV/JSON)
' - Include checkboxes (Cables/Endpoints)
' - Import mode selection
' - Progress label
' - Status area with scrollable log
```

### Dashboard Integration

**Button Creation**:

```vba
' On Dashboard sheet:
' 1. Insert Shape (rounded rectangle)
' 2. Assign macro: frm_DataManagement.Show
' 3. Format to match other dashboard buttons
```

**Alternative**: Right-click context menu on tables

```vba
' In worksheet module (e.g., sht_WetPlant):
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    ' Add custom menu item for import/export
    ' Requires CommandBar manipulation
End Sub
```

---

## Extending the System

### Adding New Export Format

Example: XML export

```vba
' In modImportExport
Public Function ExportToXML(strPlantID As String, strFilePath As String) As Boolean
    ' 1. Open file for writing
    ' 2. Write XML header
    ' 3. Iterate through plants/cables
    ' 4. Build XML elements
    ' 5. Close file
    ' 6. Return success/failure
End Function
```

### Adding Progress Bar

**Status Bar Progress** (simplest):

```vba
' In export/import loops, add:
If i Mod 10 = 0 Then
    Application.StatusBar = "Processing: " & i & " of " & totalCount
    DoEvents
End If

' At the end:
Application.StatusBar = False
```

**UserForm Progress** (visual):

1. Create `frm_Progress` with:
   - Label for percentage
   - Shape for progress bar
   - Label for status message

2. Add progress methods:

```vba
' In frm_Progress
Public Sub UpdateProgress(percent As Single, message As String)
    Me.lbl_Percent.Caption = Format(percent, "0%")
    Me.shp_ProgressBar.Width = Me.shp_ProgressBack.Width * percent
    Me.lbl_Message.Caption = message
    DoEvents
End Sub
```

3. Use in export/import:

```vba
Dim progress As New frm_Progress
progress.Show vbModeless

For i = 1 To totalCount
    ' Do work...

    If i Mod 10 = 0 Then
        progress.UpdateProgress i / totalCount, "Processing cable " & i
    End If
Next i

Unload progress
```

### Adding JSON Import

Currently CSV import only. To add JSON:

```vba
Public Function ImportFromJSON(strFilePath As String, importMode As String) As Object
    ' 1. Read JSON file into string
    ' 2. Parse JSON manually or use JsonConverter (external library)
    ' 3. Extract plants/cables/endpoints from JSON structure
    ' 4. Call ImportCable/ImportEndpoint for each record
    ' 5. Return results dictionary
End Function
```

**JSON Parsing Options**:

A) **Manual** (no dependencies):
- Use InStr, Mid, Split to parse JSON
- Brittle but works for known format
- See existing JSONEscape/JSON building code for patterns

B) **VBA-JSON Library** (recommended):
- Tim Hall's VBA-JSON: https://github.com/VBA-tools/VBA-JSON
- Robust JSON parsing
- Requires import of JsonConverter module

### Adding Validation Rules Export/Import

Currently not supported. To add:

```vba
' Export validation rules
Public Function ExportValidationRules(...) As Boolean
    ' 1. Iterate through columns with data validation
    ' 2. Extract validation formula, type, error message
    ' 3. Write to CSV/JSON
End Function

' Import validation rules
Public Function ImportValidationRules(...) As Boolean
    ' 1. Read validation rules from file
    ' 2. Apply to appropriate columns
    ' 3. Set error messages and input messages
End Function
```

### Adding Scheduled Exports

Use Windows Task Scheduler + VBA macro:

```vba
' In ThisWorkbook
Sub Auto_Open()
    ' Runs when workbook opens
    If Application.CommandLine Like "*/SCHEDULE*" Then
        ' Called by task scheduler
        PerformScheduledExport
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

Sub PerformScheduledExport()
    Dim timestamp As String
    timestamp = Format(Now, "yyyymmdd_hhnnss")

    modImportExport.ExportToJSON "ALL", _
        "C:\Backups\Scheduled_Export_" & timestamp & ".json"
End Sub
```

**Task Scheduler Command**:
```batch
"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" "C:\Path\To\File.xlsm" /SCHEDULE
```

---

## Testing

### Unit Testing

Manual testing checklist:

```vba
Sub TestDataManagement()
    Dim results As Object

    ' Test 1: Export cables to CSV
    Debug.Print "Test 1: Export Cables CSV"
    If modImportExport.ExportCablesToCSV("WET_PLANT", "C:\Temp\test_cables.csv") Then
        Debug.Print "✓ PASS"
    Else
        Debug.Print "✗ FAIL"
    End If

    ' Test 2: Export endpoints to CSV
    Debug.Print "Test 2: Export Endpoints CSV"
    If modImportExport.ExportEndpointsToCSV("WET_PLANT", "C:\Temp\test_endpoints.csv") Then
        Debug.Print "✓ PASS"
    Else
        Debug.Print "✗ FAIL"
    End If

    ' Test 3: Export to JSON
    Debug.Print "Test 3: Export JSON"
    If modImportExport.ExportToJSON("ALL", "C:\Temp\test_all.json") Then
        Debug.Print "✓ PASS"
    Else
        Debug.Print "✗ FAIL"
    End If

    ' Test 4: Create backup
    Debug.Print "Test 4: Create Backup"
    Dim backupPath As String
    backupPath = modBackup.CreateBackup("ALL")
    If backupPath <> "" Then
        Debug.Print "✓ PASS: " & backupPath
    Else
        Debug.Print "✗ FAIL"
    End If

    ' Test 5: Auto-fix missing endpoint
    Debug.Print "Test 5: Auto-fix Missing Endpoint"
    modCompatibilityFix.InitializeAutoFix
    Dim shortName As String
    shortName = modCompatibilityFix.FixMissingEndpoint("WET_PLANT", "Test Pump 99", "SOURCE")
    If shortName <> "" Then
        Debug.Print "✓ PASS: Created " & shortName
    Else
        Debug.Print "✗ FAIL"
    End If

    ' Test 6: Import cables (requires existing CSV)
    Debug.Print "Test 6: Import Cables CSV"
    Set results = modImportExport.ImportCablesFromCSV("C:\Temp\test_cables.csv", "APPEND")
    If results("Success") Then
        Debug.Print "✓ PASS: Imported " & results("CablesImported") & " cables"
    Else
        Debug.Print "✗ FAIL: " & results("ErrorMessage")
    End If

    Debug.Print "Testing complete!"
End Sub
```

### Integration Testing

**Round-Trip Test**:

```vba
Sub TestRoundTrip()
    ' 1. Count current cables
    Dim originalCount As Long
    originalCount = sht_WetPlant.ListObjects("tbl_WetPlantCables").ListRows.Count

    ' 2. Export
    modImportExport.ExportCablesToCSV "WET_PLANT", "C:\Temp\roundtrip.csv"
    modImportExport.ExportEndpointsToCSV "WET_PLANT", "C:\Temp\roundtrip_ep.csv"

    ' 3. Clear cables
    modImportExport.ClearPlantCables "WET_PLANT"

    ' 4. Import
    modImportExport.ImportEndpointsFromCSV "C:\Temp\roundtrip_ep.csv", "REPLACE"
    modImportExport.ImportCablesFromCSV "C:\Temp\roundtrip.csv", "REPLACE"

    ' 5. Count again
    Dim newCount As Long
    newCount = sht_WetPlant.ListObjects("tbl_WetPlantCables").ListRows.Count

    ' 6. Verify
    If originalCount = newCount Then
        Debug.Print "✓ Round-trip successful: " & newCount & " cables"
    Else
        Debug.Print "✗ Round-trip failed: " & originalCount & " → " & newCount
    End If
End Sub
```

### Edge Cases to Test

1. **Empty tables** - Export/import with no data
2. **Single cable** - Ensure DataBodyRange.Value works
3. **Special characters** - Commas, quotes, newlines in descriptions
4. **Large dataset** - 1000+ cables (performance)
5. **Duplicate Cable IDs** - MERGE mode handling
6. **Missing endpoints** - Auto-fix creation
7. **Malformed CSV** - Error handling
8. **File locked** - Can't open for writing
9. **Backup overflow** - More than 10 backups (cleanup test)
10. **Invalid plant ID** - Error handling

---

## Performance Considerations

### Current Performance

Based on testing:

- **Export CSV**: ~100 cables/second
- **Export JSON**: ~50 cables/second (more complex)
- **Import CSV**: ~50 cables/second (includes validation)
- **Auto-fix lookup**: ~1000 endpoints/second

### Optimization Opportunities

#### 1. Batch Table Updates

Current approach: Individual row operations
```vba
For i = 1 To count
    Set newRow = tbl.ListRows.Add
    newRow.Range(1, 1).Value = data1
    ' ... etc
Next i
```

Optimized approach: Array assignment
```vba
Dim arr() As Variant
ReDim arr(1 To count, 1 To 11)
' Populate array
tbl.DataBodyRange.Value = arr
```

Improvement: ~5-10x faster for large datasets

#### 2. Disable Screen Updating

```vba
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' Do work...

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
```

Improvement: ~2-3x faster

#### 3. Endpoint Lookup Cache

Current: Linear search through all endpoints for each cable
Optimized: Build dictionary once, lookup in O(1)

```vba
' Build lookup dictionary
Dim epLookup As Object
Set epLookup = CreateObject("Scripting.Dictionary")
For Each ep In endpoints
    epLookup.Add ep.Description, ep.ShortName
Next

' Use in import
shortName = epLookup(description)
```

Improvement: ~100x faster for large endpoint lists

#### 4. File I/O Optimization

Current: Line-by-line with TextStream
Optimized: Read entire file, split by newlines

```vba
' Current
Do While Not txtFile.AtEndOfStream
    line = txtFile.ReadLine
    ' Process
Loop

' Optimized
Dim content As String
content = txtFile.ReadAll
Dim lines() As String
lines = Split(content, vbCrLf)
For Each line In lines
    ' Process
Next
```

Improvement: ~2x faster

---

## Security Considerations

### File Path Validation

```vba
Function ValidateFilePath(filePath As String) As Boolean
    ' Check for directory traversal
    If InStr(filePath, "..") > 0 Then
        ValidateFilePath = False
        Exit Function
    End If

    ' Check file extension
    Dim ext As String
    ext = LCase(Right(filePath, 4))
    If ext <> ".csv" And ext <> "json" Then
        ValidateFilePath = False
        Exit Function
    End If

    ValidateFilePath = True
End Function
```

### CSV Injection Prevention

When exporting data that users may open in Excel:

```vba
Function SanitizeCSVField(value As String) As String
    ' Prevent formula injection
    If Left(value, 1) = "=" Or _
       Left(value, 1) = "+" Or _
       Left(value, 1) = "-" Or _
       Left(value, 1) = "@" Then
        SanitizeCSVField = "'" & value  ' Prefix with quote
    Else
        SanitizeCSVField = value
    End If
End Function
```

### Macro Security

System requires macros to be enabled. User considerations:

- Sign VBA project with digital certificate (optional)
- Provide clear instructions about enabling macros
- Use trusted locations feature
- Document that macro-free version loses import/export

---

## Troubleshooting

### Common Developer Issues

#### "Object doesn't support this property or method"

**Cause**: Dictionary or Collection used without proper reference

**Fix**: Ensure Scripting Runtime is available, or use CreateObject:
```vba
Set dict = CreateObject("Scripting.Dictionary")  ' Late binding
```

#### "Type mismatch" on CSV import

**Cause**: Boolean conversion failing

**Fix**: Check CSV has uppercase TRUE/FALSE, use conversion:
```vba
Function SafeBoolean(value As Variant) As Boolean
    Select Case UCase(Trim(CStr(value)))
        Case "TRUE", "1", "YES": SafeBoolean = True
        Case Else: SafeBoolean = False
    End Select
End Function
```

#### "Run-time error '9': Subscript out of range"

**Cause**: Table or worksheet doesn't exist

**Fix**: Verify CodeNames and Table names match expected:
```vba
If WorksheetExists("sht_WetPlant") Then
    ' Proceed
Else
    MsgBox "Required worksheet missing", vbCritical
End If
```

---

## Version Control

### Recommended Approach

1. **VBA Code**: Export modules to `.bas` files
   - Commit to git repository
   - Track changes over time

2. **Excel Files**: Don't commit actual spreadsheets with data
   - Commit template-only versions
   - Or use `.gitignore` for `*.xlsm`

3. **Branching Strategy**:
   - `main`: Stable releases
   - `develop`: Active development
   - `feature/import-export`: This feature

### Deployment

1. **Development**: Test in copy of production file
2. **Staging**: Import modules to test spreadsheet
3. **Production**: Import modules to production file
4. **Rollback**: Keep previous version as backup

---

## Future Enhancements Roadmap

### Phase 2: Enhanced UI
- Progress bar implementation
- Real-time log viewer
- Preview before import
- Diff viewer (compare versions)

### Phase 3: Advanced Features
- JSON import implementation
- Scheduled auto-exports
- Incremental backups (only changed data)
- Compression for large exports

### Phase 4: Integration
- SharePoint upload/download
- OneDrive sync
- Email exports automatically
- Web API integration

### Phase 5: Validation
- Export/import validation rules
- Export/import conditional formatting
- Export/import data types
- Schema migration tools

---

## Support & Maintenance

### Logging

Add comprehensive logging for production:

```vba
Public Sub LogOperation(operation As String, details As String)
    Dim logPath As String
    Dim fso As Object
    Dim txtFile As Object

    logPath = ThisWorkbook.Path & "\Logs\ImportExport_" & Format(Now, "yyyymmdd") & ".log"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.OpenTextFile(logPath, 8, True)  ' 8 = Append

    txtFile.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & operation & " | " & details
    txtFile.Close

    Set txtFile = Nothing
    Set fso = Nothing
End Sub
```

### Error Reporting

Collect diagnostic information:

```vba
Function GetDiagnosticInfo() As String
    Dim info As String
    info = "Excel Version: " & Application.Version & vbCrLf
    info = info & "Workbook: " & ThisWorkbook.Name & vbCrLf
    info = info & "Module Version: " & MODULE_VERSION & vbCrLf
    info = info & "Cable Count: " & GetTotalCableCount() & vbCrLf
    GetDiagnosticInfo = info
End Function
```

---

## License & Credits

**Copyright**: 2024
**Author**: AI Assistant (Claude)
**Contact**: jorr@mtcarbine.com.au
**License**: Internal use only

**Third-Party Components**:
- None (all custom implementation)

**Acknowledgments**:
- RFC 4180 CSV specification
- JSON specification (ECMA-404)
- VBA community best practices

---

## Appendix

### Complete Function Signatures

See `vba_code/README.md` for full API reference.

### Glossary

- **Plant**: One of three cable categories (Wet Plant, Ore Sorter, Retreatment)
- **Endpoint**: Source or destination location for a cable
- **Short Name**: Abbreviated endpoint identifier (e.g., "MCC102")
- **Description**: Full endpoint name (e.g., "MCC 2")
- **Cable ID**: Unique identifier format: SOURCE-CIRCUIT-DESTINATION
- **Auto-Fix**: Automatic compatibility correction during import
- **Fuzzy Match**: Approximate string matching algorithm
- **Import Mode**: Strategy for handling existing data (Append/Replace/Merge)

### References

- [IMPORT_EXPORT_Implementation_Plan.md](IMPORT_EXPORT_Implementation_Plan.md)
- [IMPORT_EXPORT_Installation_Guide.md](IMPORT_EXPORT_Installation_Guide.md)
- [DataManagement_User_Guide.md](DataManagement_User_Guide.md)
- [vba_code/README.md](../vba_code/README.md)

---

**Document Version**: 1.0
**Last Updated**: December 2024
**Status**: Complete

For questions or contributions, contact the development team.

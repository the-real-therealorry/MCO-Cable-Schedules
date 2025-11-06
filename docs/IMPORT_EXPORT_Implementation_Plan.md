# Cable Import/Export Feature - Implementation Plan

## Overview

A comprehensive data migration system to transfer cables and endpoints between spreadsheet versions, with automatic compatibility fixing.

---

## Requirements Summary

1. **File Formats**: CSV and JSON
2. **Export Data**: Cables AND Endpoints
3. **Export Scope**: All cables (all plants) OR individual plant
4. **Import Mode**: Ask user (Append / Replace / Merge)
5. **Validation**: Import anyway, show warnings after completion
6. **UI**: Dedicated "Data Management" form, launched from Dashboard button
7. **Error Handling**: Summary report + log file
8. **Compatibility**: Auto-fix incompatibilities on the fly

---

## File Format Specifications

### CSV Format

**Cables Export**: `Cables_Export_YYYYMMDD_HHMMSS.csv`

```csv
Version,Plant,Scheduled,IDAttached,CableID,Source,Destination,CoreSize,EarthSize,CoreConfig,InsulationType,CableType,CableLength
1.0,WET_PLANT,FALSE,TRUE,CV102-C1001-CV103,MCC 2,Conveyor belt 3,2.5mm²,1.5mm²,4C + E,HFI-90-TP,Black Circular SWA,45
```

**Endpoints Export**: `Endpoints_Export_YYYYMMDD_HHMMSS.csv`

```csv
Version,Plant,ShortName,Description
1.0,WET_PLANT,MCC102,MCC 2
1.0,WET_PLANT,CV103,Conveyor belt 3
```

**Header Row**:
- Version: Spreadsheet version (for compatibility detection)
- Plant: WET_PLANT, ORE_SORTER, or RETREATMENT

### JSON Format

**Combined Export**: `CableRegister_Export_YYYYMMDD_HHMMSS.json`

```json
{
  "version": "1.0",
  "exportDate": "2024-12-21T10:30:00Z",
  "sourceFile": "Power Cable Register Rev0.xlsm",
  "plants": {
    "WET_PLANT": {
      "endpoints": [
        {
          "shortName": "MCC102",
          "description": "MCC 2"
        },
        {
          "shortName": "CV103",
          "description": "Conveyor belt 3"
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
    },
    "ORE_SORTER": {
      "endpoints": [],
      "cables": []
    },
    "RETREATMENT": {
      "endpoints": [],
      "cables": []
    }
  },
  "metadata": {
    "totalCables": 1,
    "totalEndpoints": 2,
    "exportedBy": "User",
    "exportType": "ALL_PLANTS"
  }
}
```

---

## Auto-Fix Compatibility System

### Version Detection

When importing, detect version mismatches and apply fixes:

```vba
Function DetectVersionCompatibility(importVersion As String, currentVersion As String) As String
    ' Returns: "COMPATIBLE", "NEEDS_FIX", "INCOMPATIBLE"
    ' Auto-fix handles: minor version differences, column renames, endpoint changes
End Function
```

### Auto-Fix Strategies

#### 1. Missing Endpoints

**Problem**: Imported cable references "MCC 2" but endpoint doesn't exist in new version

**Auto-Fix Options** (in order of preference):

a) **Exact Short Name Match**
   - Check if short name exists in new version (MCC102)
   - If yes, update description to match new version

b) **Fuzzy Description Match**
   - Try variations: "MCC 2" → "MCC 02" → "MCC2"
   - Check for close matches (Levenshtein distance)

c) **Create Missing Endpoint**
   - Auto-create the endpoint in the endpoints table
   - Mark as "Imported - Needs Review"
   - Add to compatibility report

d) **Leave as-is with Warning**
   - Import cable with original description
   - Flag in warning report
   - User can fix manually later

**Implementation**:
```vba
Function FixMissingEndpoint(plantID As String, endpointDesc As String) As String
    ' Try exact match
    ' Try fuzzy match
    ' Create if needed
    ' Return: corrected description or original
End Function
```

#### 2. Column Mapping

**Problem**: Old version has "Insulation" but new version renamed to "InsulationType"

**Auto-Fix**:
- Maintain column mapping dictionary
- Auto-map old → new column names
- Handle missing columns with defaults

```vba
Function GetColumnMapping(oldVersion As String, newVersion As String) As Dictionary
    ' Returns mapping: oldColumnName → newColumnName
    ' E.g., "Insulation" → "InsulationType"
End Function
```

#### 3. Data Type Conversion

**Problem**: Old version stored length as text "45m", new version expects numeric "45"

**Auto-Fix**:
- Strip units (m, mm, etc.)
- Convert text to numbers
- Handle null/empty gracefully

```vba
Function ConvertDataType(value As Variant, targetType As String) As Variant
    ' Auto-convert with error handling
End Function
```

#### 4. Validation Rule Changes

**Problem**: New version has stricter Cable ID format validation

**Auto-Fix**:
- Attempt to reformat to new pattern
- If impossible, import as-is and flag
- Add to "Needs Review" list

#### 5. Plant ID Changes

**Problem**: Old version used "WETPLANT", new version uses "WET_PLANT"

**Auto-Fix**:
- Normalize plant identifiers
- Map old → new naming

```vba
Function NormalizePlantID(oldPlantID As String) As String
    ' Handle variations in plant naming
End Function
```

### Compatibility Report

After import with auto-fixes, generate detailed report:

```
=== IMPORT COMPATIBILITY REPORT ===
Date: 2024-12-21 10:30:00
Source: Cables_Export_20241220.csv
Target: Power Cable Register Rev1.xlsm

SUMMARY:
✓ 45 cables imported successfully
⚠ 3 cables auto-fixed
✗ 0 cables failed

AUTO-FIXES APPLIED:
1. Missing Endpoint: "MCC 2" → Created new endpoint MCC102
2. Column Rename: "Insulation" → "InsulationType" (3 cables)
3. Data Format: Converted "45m" → "45" (2 cables)

WARNINGS:
- Cable CV102-C1001-CV103: Endpoint "OLD_PUMP" not found, created as new
- Cable OS201-C2001-OS202: Length format unusual, please verify

RECOMMENDATIONS:
1. Review newly created endpoints in Data sheet
2. Verify 3 cables with format conversions
3. Backup created: Import_Backup_20241221_103000.json
```

---

## Module Structure

### New VBA Modules

```
modImportExport.bas
├── ExportCables()
├── ExportEndpoints()
├── ExportToCSV()
├── ExportToJSON()
├── ImportCables()
├── ImportFromCSV()
├── ImportFromJSON()
└── GenerateCompatibilityReport()

modCompatibilityFix.bas
├── DetectVersionCompatibility()
├── FixMissingEndpoint()
├── GetColumnMapping()
├── ConvertDataType()
├── NormalizePlantID()
└── ApplyAutoFixes()

frm_DataManagement.frm
├── UI for export/import options
├── File browser dialogs
├── Progress indicators
└── Results display
```

---

## User Interface Design

### Dashboard Button

**Add to Dashboard sheet**:
- Shape/Button: "Data Management" or "Import/Export"
- Position: Near other management functions
- Opens: frm_DataManagement

### Data Management Form

**Layout**:

```
┌─────────────────────────────────────────────────┐
│  Cable Register - Data Management               │
├─────────────────────────────────────────────────┤
│                                                  │
│  ┌─── EXPORT ─────────────────────────────────┐ │
│  │                                             │ │
│  │  Export Scope:                              │ │
│  │  ○ All Plants                               │ │
│  │  ○ Wet Plant Only                           │ │
│  │  ○ Ore Sorter Only                          │ │
│  │  ○ Retreatment Only                         │ │
│  │                                             │ │
│  │  Format:                                    │ │
│  │  ☑ CSV  ☑ JSON                              │ │
│  │                                             │ │
│  │  Include:                                   │ │
│  │  ☑ Cables  ☑ Endpoints                      │ │
│  │                                             │ │
│  │  [Export to File...]                        │ │
│  └─────────────────────────────────────────────┘ │
│                                                  │
│  ┌─── IMPORT ─────────────────────────────────┐ │
│  │                                             │ │
│  │  Import Mode:                               │ │
│  │  ○ Append (add to existing)                 │ │
│  │  ○ Replace (clear and import)               │ │
│  │  ○ Merge (update by Cable ID)               │ │
│  │                                             │ │
│  │  Auto-Fix Compatibility:                    │ │
│  │  ☑ Enabled (recommended)                    │ │
│  │                                             │ │
│  │  [Import from File...]                      │ │
│  └─────────────────────────────────────────────┘ │
│                                                  │
│  ┌─── LAST OPERATION ─────────────────────────┐ │
│  │ Status: Ready                               │ │
│  │                                             │ │
│  └─────────────────────────────────────────────┘ │
│                                                  │
│               [Close]                            │
└─────────────────────────────────────────────────┘
```

---

## Implementation Steps

### Phase 1: Basic Export (Week 1)

1. Create `modImportExport` module
2. Implement CSV export for cables
3. Implement CSV export for endpoints
4. Add Dashboard button
5. Create basic frm_DataManagement
6. Test export with current data

**Deliverable**: Working CSV export

### Phase 2: Basic Import (Week 1-2)

1. Implement CSV import for cables
2. Implement CSV import for endpoints
3. Add import mode selection (Append/Replace/Merge)
4. Add basic validation
5. Test round-trip (export then import)

**Deliverable**: Working CSV import/export

### Phase 3: JSON Support (Week 2)

1. Add JSON export functionality
2. Add JSON import functionality
3. Handle nested structure
4. Test combined export/import

**Deliverable**: Working JSON import/export

### Phase 4: Auto-Fix System (Week 2-3)

1. Create `modCompatibilityFix` module
2. Implement version detection
3. Implement missing endpoint auto-fix
4. Implement column mapping
5. Implement data type conversion
6. Test with intentional incompatibilities

**Deliverable**: Working auto-fix system

### Phase 5: Reporting & Polish (Week 3)

1. Implement compatibility report generation
2. Add log file creation
3. Create backup before import
4. Add progress indicators
5. Improve error messages
6. Final testing

**Deliverable**: Production-ready feature

---

## Error Handling Strategy

### Three-Level Approach

**1. Auto-Fix (Silent)**
- Minor issues fixed automatically
- Logged but no user interruption
- Examples: column renames, whitespace trimming

**2. Warning (Notification)**
- Fixed but user should review
- Shown in summary report
- Examples: missing endpoints created, fuzzy matches

**3. Error (Block/Skip)**
- Cannot be auto-fixed
- Skip item or fail import
- Examples: completely invalid data, corrupt file

### Rollback Capability

Before any import:
1. Create backup file: `Import_Backup_YYYYMMDD_HHMMSS.json`
2. Store current state
3. If import fails catastrophically, offer rollback

```vba
Function RollbackImport(backupFile As String) As Boolean
    ' Restore from backup
    ' Show confirmation
End Function
```

---

## Testing Plan

### Test Scenarios

**1. Basic Round-Trip**
- Export all cables → Import to fresh file → Verify identical

**2. Version Upgrade Simulation**
- Export from v1.0 → Rename columns in v2.0 → Import → Verify auto-fix

**3. Missing Endpoints**
- Export cables → Delete endpoints → Import → Verify auto-creation

**4. Partial Import**
- Export all → Import only Wet Plant → Verify scope

**5. Format Conversion**
- Export both CSV and JSON → Import both → Verify identical results

**6. Error Conditions**
- Corrupt file
- Wrong file type
- Empty file
- Duplicate cable IDs (Merge mode)

### Acceptance Criteria

✅ Export all cables to CSV and JSON
✅ Import from CSV and JSON
✅ Auto-fix missing endpoints
✅ Handle column renames
✅ Generate compatibility report
✅ Create import log file
✅ Support Append/Replace/Merge modes
✅ Show progress during long operations
✅ Create backups before import
✅ Rollback capability

---

## File Structure

```
Project Root/
├── Power Cable Register Rev0.xlsm
├── Exports/
│   ├── Cables_Export_20241221_103000.csv
│   ├── Endpoints_Export_20241221_103000.csv
│   └── CableRegister_Export_20241221_103000.json
├── Imports/
│   └── (user's import files)
├── Backups/
│   └── Import_Backup_20241221_103000.json
└── Logs/
    └── Import_Log_20241221_103000.txt
```

---

## Future Enhancements

### Phase 6+ (Optional)

1. **Cloud Integration**
   - Export to SharePoint
   - Import from SharePoint
   - Sync with central database

2. **Advanced Filters**
   - Export cables by date range
   - Export by cable type
   - Custom filters

3. **Template Management**
   - Save export templates
   - Scheduled exports

4. **Validation Rules Export**
   - Export data validation rules
   - Import and apply to new version

5. **Diff Viewer**
   - Compare two versions
   - Show what changed
   - Selective import

---

## Questions Before Implementation

1. **Version Numbering**: How should we version the spreadsheets? (e.g., 1.0, 1.1, 2.0)

2. **Default Export Location**: Should exports go to:
   - Same folder as workbook?
   - Documents folder?
   - User chooses each time?

3. **Auto-Create Endpoints**: When auto-creating missing endpoints, should they:
   - Go into the same plant's endpoint table?
   - Get flagged somehow for review?

4. **Backup Retention**: How long to keep backup files?
   - Keep all backups?
   - Delete after X days?
   - Keep last N backups?

5. **Large Datasets**: If you have 1000+ cables, should we:
   - Show progress bar?
   - Process in batches?
   - Background processing?

Let me know your preferences and I'll start implementing!
